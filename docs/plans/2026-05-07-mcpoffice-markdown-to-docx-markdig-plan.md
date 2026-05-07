# Markdown→DOCX via Markdig — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Replace the lossy `MarkdownToDocxGenerator` v1.2.0 path with a Markdig-based AST walker that emits DevExpress `Document` API calls directly, fixing inline-code stripping, GFM table flattening, and bold-run loss.

**Architecture:** New `MarkdownToDocxConverter` class parses markdown with Markdig (CommonMark + `UsePipeTables()`), walks the AST, and populates a DevExpress `Document` via `Paragraphs`/`Tables`/`Images` API. Stateless, single static `Apply(Document, string, string?)` entry point. `WordDocumentService.CreateFromMarkdown`/`AppendMarkdown` and the markdown branch of `Convert` become thin wrappers. Old converter package + post-process regex helpers are deleted at the end.

**Tech Stack:** .NET 9 · Markdig · DevExpress.Document.Processor (`RichEditDocumentServer`) · xUnit + FluentAssertions.

**Reference design:** `docs/plans/2026-05-07-mcpoffice-markdown-to-docx-markdig-design.md` — single source of truth for mapping rules, error policy, and out-of-scope items. Read it before starting.

---

## Conventions used in this plan

- All paths are relative to `C:\Projects\mcpOffice\` (the repo root).
- "Run tests" means `dotnet test --nologo --logger "console;verbosity=normal"` from the repo root unless a more specific filter is given.
- "Build clean" means `dotnet build --nologo` returns 0 warnings, 0 errors.
- Each task ends with a Conventional Commits commit. Branch is `feat/markdown-to-docx-markdig` (already created).
- New converter unit tests construct a real `RichEditDocumentServer` in-memory — no mocking, no file I/O, no DevExpress license worries (matches the existing `OutlineTests` pattern).
- Exact DevExpress API calls are specified where well-known. Where the spike risks from the design apply (character shading, alignment), the task notes the fallback.

---

# Phase A — Bootstrap

### Task 1: Add Markdig package and converter skeleton

**Files:**
- Modify: `src/mcpOffice/mcpOffice.csproj`
- Create: `src/mcpOffice/Services/Word/MarkdownToDocxConverter.cs`
- Create: `tests/mcpOffice.Tests/Word/MarkdownToDocxConverterTests.cs`

**Step 1: Add Markdig**
```bash
dotnet add src/mcpOffice package Markdig
```
Expected: package added, restore succeeds, build still clean.

**Step 2: Create converter skeleton**
```csharp
// src/mcpOffice/Services/Word/MarkdownToDocxConverter.cs
using DevExpress.XtraRichEdit.API.Native;
using Markdig;
using Markdig.Syntax;

namespace McpOffice.Services.Word;

internal static class MarkdownToDocxConverter
{
    private static readonly MarkdownPipeline Pipeline =
        new MarkdownPipelineBuilder().UsePipeTables().Build();

    public static void Apply(Document document, string markdown, string? baseDirectory)
    {
        var ast = Markdown.Parse(markdown ?? string.Empty, Pipeline);
        // Block dispatch added in subsequent tasks.
        _ = ast;
        _ = document;
        _ = baseDirectory;
    }
}
```

**Step 3: Write skeleton smoke test** — proves wiring compiles
```csharp
// tests/mcpOffice.Tests/Word/MarkdownToDocxConverterTests.cs
using DevExpress.XtraRichEdit;
using FluentAssertions;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class MarkdownToDocxConverterTests
{
    private static DevExpress.XtraRichEdit.API.Native.Document NewDocument(out RichEditDocumentServer server)
    {
        server = new RichEditDocumentServer();
        return server.Document;
    }

    [Fact]
    public void Apply_with_empty_markdown_does_not_throw()
    {
        var doc = NewDocument(out var server);
        try
        {
            var act = () => MarkdownToDocxConverter.Apply(doc, string.Empty, null);
            act.Should().NotThrow();
        }
        finally { server.Dispose(); }
    }
}
```

**Step 4: Build + run the new test**
```bash
dotnet build --nologo && dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~MarkdownToDocxConverterTests --nologo
```
Expected: build clean, 1 passed.

**Step 5: Commit**
```bash
git add src/mcpOffice/mcpOffice.csproj src/mcpOffice/Services/Word/MarkdownToDocxConverter.cs tests/mcpOffice.Tests/Word/MarkdownToDocxConverterTests.cs
git commit -m "chore: add Markdig + MarkdownToDocxConverter skeleton"
```

---

# Phase B — Block-level mapping (TDD per rule)

> **Pattern for every task in this phase:**
> 1. Add a failing test to `MarkdownToDocxConverterTests`.
> 2. Run it — fail (`Apply` no-ops or hits an unhandled block type).
> 3. Add the dispatch + helper in `MarkdownToDocxConverter` until the test passes.
> 4. Run all converter tests — every prior one still green.
> 5. Commit.

The test helper `NewDocument(out var server)` from Task 1 is reused. Where DevExpress APIs are unobvious, the task gives the exact call.

### Task 2: Empty markdown produces empty document

**Step 1: Test**
```csharp
[Fact]
public void Empty_markdown_produces_empty_document()
{
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, "", null);
    server.Document.Paragraphs.Count.Should().Be(1);
    server.Document.GetText(server.Document.Range).Trim().Should().BeEmpty();
}
```

**Step 2:** Already passes (skeleton no-ops). This test is the floor we never break.

**Step 3: Commit**
```bash
git commit -am "test: empty markdown produces empty document"
```

---

### Task 3: Plain paragraphs

**Step 1: Test**
```csharp
[Fact]
public void Plain_paragraphs_become_paragraphs_with_literal_text()
{
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, "hello world\n\nsecond para", null);
    var text = server.Document.GetText(server.Document.Range).Trim();
    text.Should().Contain("hello world").And.Contain("second para");
    server.Document.Paragraphs.Count.Should().BeGreaterThanOrEqualTo(2);
}
```

**Step 2: Implement** — the foundation for all subsequent inlines

Add the block dispatcher and a literal-only inline walker. In `MarkdownToDocxConverter.Apply`:

```csharp
public static void Apply(Document document, string markdown, string? baseDirectory)
{
    var ast = Markdown.Parse(markdown ?? string.Empty, Pipeline);
    var ctx = new ConversionContext(document, baseDirectory);
    foreach (var block in ast)
        WriteBlock(ctx, block);
}

private sealed record ConversionContext(Document Document, string? BaseDirectory);

private static void WriteBlock(ConversionContext ctx, Block block)
{
    switch (block)
    {
        case ParagraphBlock p:
            WriteParagraph(ctx, p);
            break;
        // Other cases added in later tasks.
        default:
            // Skipped silently for now; Task 21 adds the Serilog warning.
            break;
    }
}

private static void WriteParagraph(ConversionContext ctx, ParagraphBlock block)
{
    var para = AppendNewParagraph(ctx);
    if (block.Inline is null) return;
    foreach (var inline in block.Inline)
        WriteInline(ctx, para, inline);
}

private static Paragraph AppendNewParagraph(ConversionContext ctx)
{
    var doc = ctx.Document;
    var pos = doc.Range.End;
    doc.InsertParagraph(pos);
    return doc.Paragraphs[^1];
}

private static void WriteInline(ConversionContext ctx, Paragraph para, Markdig.Syntax.Inlines.Inline inline)
{
    switch (inline)
    {
        case Markdig.Syntax.Inlines.LiteralInline lit:
            ctx.Document.InsertText(para.Range.End, lit.Content.ToString());
            break;
        // Bold/italic/code added in Phase C.
    }
}
```

**Step 3: Run**
```bash
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~MarkdownToDocxConverterTests --nologo
```
Expected: 3 passed.

**Step 4: Commit**
```bash
git commit -am "feat(markdown): paragraph + literal inline"
```

---

### Task 4: Headings 1–6

**Step 1: Test**
```csharp
[Fact]
public void Headings_1_through_6_get_correct_paragraph_style()
{
    var md = "# h1\n\n## h2\n\n### h3\n\n#### h4\n\n##### h5\n\n###### h6";
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);

    var headingParas = server.Document.Paragraphs
        .Where(p => p.Style?.Name?.StartsWith("Heading ") == true)
        .ToList();

    headingParas.Should().HaveCount(6);
    for (int i = 0; i < 6; i++)
        headingParas[i].Style!.Name.Should().Be($"Heading {i + 1}");
}
```

**Step 2: Implement** — add `HeadingBlock` case

In `WriteBlock`:
```csharp
case HeadingBlock h:
    WriteHeading(ctx, h);
    break;
```

```csharp
private static void WriteHeading(ConversionContext ctx, HeadingBlock block)
{
    var styleName = $"Heading {Math.Clamp(block.Level, 1, 6)}";
    EnsureParagraphStyle(ctx.Document, styleName);
    var para = AppendNewParagraph(ctx);
    para.Style = ctx.Document.ParagraphStyles[styleName];
    if (block.Inline is null) return;
    foreach (var inline in block.Inline)
        WriteInline(ctx, para, inline);
}

private static void EnsureParagraphStyle(Document doc, string styleName)
{
    if (doc.ParagraphStyles[styleName] is not null) return;
    var s = doc.ParagraphStyles.CreateNew();
    s.Name = styleName;
    doc.ParagraphStyles.Add(s);
}
```

(Helper duplicates the one in `WordDocumentService`; we'll de-dupe at the end of Phase D once the migration is complete.)

**Step 3: Run, commit**
```bash
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~MarkdownToDocxConverterTests --nologo
git commit -am "feat(markdown): headings 1-6 -> Heading {N} style"
```

---

### Task 5: Unordered + ordered lists (flat)

**Step 1: Test**
```csharp
[Fact]
public void Unordered_list_produces_bulleted_paragraphs()
{
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, "- a\n- b\n- c", null);
    var listParas = server.Document.Paragraphs
        .Where(p => p.ListIndex >= 0)
        .ToList();
    listParas.Should().HaveCount(3);
    listParas.All(p => p.ListLevel == 0).Should().BeTrue();
}

[Fact]
public void Ordered_list_produces_numbered_paragraphs()
{
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, "1. one\n2. two", null);
    var listParas = server.Document.Paragraphs
        .Where(p => p.ListIndex >= 0)
        .ToList();
    listParas.Should().HaveCount(2);
}
```

**Step 2: Implement** — add `ListBlock` case + helper. DevExpress lists use `Document.NumberingLists.Add(NumberingType.Bullet)` / `NumberingType.MultiLevel` and assign list indexes to paragraphs.

```csharp
case ListBlock list:
    WriteList(ctx, list, level: 0);
    break;
```

```csharp
private static void WriteList(ConversionContext ctx, ListBlock list, int level)
{
    var numberingList = list.IsOrdered
        ? ctx.Document.NumberingLists.Add(NumberingType.MultiLevel)
        : ctx.Document.NumberingLists.Add(NumberingType.Bullet);

    foreach (var item in list.OfType<ListItemBlock>())
    {
        foreach (var sub in item)
        {
            if (sub is ParagraphBlock p)
            {
                var para = AppendNewParagraph(ctx);
                para.ListIndex = ctx.Document.NumberingLists.IndexOf(numberingList);
                para.ListLevel = level;
                if (p.Inline is not null)
                    foreach (var inline in p.Inline)
                        WriteInline(ctx, para, inline);
            }
            else if (sub is ListBlock nested)
            {
                WriteList(ctx, nested, level + 1);
            }
        }
    }
}
```

**Step 3: Run, commit**
```bash
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~MarkdownToDocxConverterTests --nologo
git commit -am "feat(markdown): ordered + unordered lists"
```

---

### Task 6: Nested list indentation

**Step 1: Test**
```csharp
[Fact]
public void Nested_list_indents_per_depth()
{
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, "- outer\n  - inner", null);
    var paras = server.Document.Paragraphs.Where(p => p.ListIndex >= 0).ToList();
    paras.Should().HaveCount(2);
    paras[0].ListLevel.Should().Be(0);
    paras[1].ListLevel.Should().Be(1);
}
```

**Step 2:** Should pass already from Task 5's recursion. If it doesn't, the issue is `ListLevel` assignment — debug accordingly.

**Step 3: Commit**
```bash
git commit -am "test: verify nested list ListLevel"
```

---

### Task 7: Blockquote

**Step 1: Test**
```csharp
[Fact]
public void Blockquote_indents_left_quarter_inch()
{
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, "> quoted text", null);
    var para = server.Document.Paragraphs.Last(p => !string.IsNullOrWhiteSpace(server.Document.GetText(p.Range)));
    para.LeftIndent.Should().BeGreaterThan(0);
}
```

**Step 2: Implement**

```csharp
case QuoteBlock q:
    WriteQuote(ctx, q);
    break;
```

```csharp
private static void WriteQuote(ConversionContext ctx, QuoteBlock block)
{
    foreach (var child in block)
    {
        if (child is ParagraphBlock p)
        {
            var para = AppendNewParagraph(ctx);
            para.LeftIndent = (float)Units.InchesToDocuments(0.25);
            if (p.Inline is null) continue;
            foreach (var inline in p.Inline)
                WriteInline(ctx, para, inline);
        }
    }
}
```

`Units.InchesToDocuments` is from `DevExpress.XtraRichEdit.API.Native`. If `LeftIndent` is integer-typed (DevExpress version), drop the cast.

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): blockquote left indent"
```

---

### Task 8: Fenced code block (and indented code block treated same)

**Step 1: Test**
```csharp
[Fact]
public void Fenced_code_block_each_line_is_monospace_paragraph()
{
    var md = "```\nfoo\nbar\n```";
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);
    var codeParas = server.Document.Paragraphs
        .Where(p => server.Document.GetText(p.Range).Trim() is "foo" or "bar")
        .ToList();
    codeParas.Should().HaveCount(2);
    foreach (var p in codeParas)
    {
        var firstRunRange = server.Document.CreateRange(p.Range.Start, 1);
        var charProps = server.Document.BeginUpdateCharacters(firstRunRange);
        try { charProps.FontName.Should().Be("Consolas"); }
        finally { server.Document.EndUpdateCharacters(charProps); }
    }
}
```

**Step 2: Implement**

```csharp
case FencedCodeBlock fenced:
    WriteCodeBlock(ctx, fenced.Lines.ToString());
    break;
case CodeBlock code when block is not FencedCodeBlock:
    WriteCodeBlock(ctx, code.Lines.ToString());
    break;
```

```csharp
private static void WriteCodeBlock(ConversionContext ctx, string text)
{
    var doc = ctx.Document;
    var lines = text.Replace("\r\n", "\n").Split('\n');
    foreach (var line in lines)
    {
        var para = AppendNewParagraph(ctx);
        para.LeftIndent = (float)Units.InchesToDocuments(0.1);
        if (line.Length > 0)
        {
            var insertedAt = doc.InsertText(para.Range.End, line);
            var charProps = doc.BeginUpdateCharacters(insertedAt);
            try
            {
                charProps.FontName = "Consolas";
                charProps.FontSize = 9f;
            }
            finally { doc.EndUpdateCharacters(charProps); }
        }
        // Paragraph shading (#F2F2F2): try ParagraphProperties.BackColor; if unavailable, leave font-only — see Risk 1 in design doc.
    }
}
```

If paragraph shading is unavailable in this DevExpress version, document the fallback in a one-line comment and move on.

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): fenced + indented code blocks (Consolas)"
```

---

### Task 9: Horizontal rule

**Step 1: Test**
```csharp
[Fact]
public void Hr_emits_paragraph_with_bottom_border()
{
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, "before\n\n---\n\nafter", null);
    var hrPara = server.Document.Paragraphs
        .FirstOrDefault(p => p.Borders?.Bottom?.LineStyle != BorderLineStyle.None);
    hrPara.Should().NotBeNull();
}
```

**Step 2: Implement**

```csharp
case ThematicBreakBlock:
    WriteHorizontalRule(ctx);
    break;
```

```csharp
private static void WriteHorizontalRule(ConversionContext ctx)
{
    var para = AppendNewParagraph(ctx);
    para.Borders.Bottom.LineStyle = BorderLineStyle.Single;
    para.Borders.Bottom.LineThickness = 0.5f;
}
```

If `Borders.Bottom` API isn't directly settable (some DevExpress versions require `BeginUpdateBorders`), use that pattern. Fallback: if borders aren't accessible, emit `"───────────"` as a literal text paragraph and leave a TODO comment.

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): thematic break -> hr paragraph"
```

---

### Task 10: GFM pipe table — basic creation + header bold/shaded

**Step 1: Test**
```csharp
[Fact]
public void Pipe_table_creates_real_table_with_bold_shaded_header()
{
    var md = "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |";
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);

    server.Document.Tables.Count.Should().Be(1);
    var table = server.Document.Tables[0];
    table.Rows.Count.Should().Be(3);                  // header + 2 data rows
    table.Rows[0].Cells.Count.Should().Be(2);

    // Header cell shading
    var headerCell = table.Rows[0].Cells[0];
    var bgColor = headerCell.BackgroundColor;
    bgColor.ToArgb().Should().Be(System.Drawing.Color.FromArgb(0xF2, 0xF2, 0xF2).ToArgb());

    // Header runs bold
    var charProps = server.Document.BeginUpdateCharacters(headerCell.Range);
    try { charProps.FontBold.Should().Be(true); }
    finally { server.Document.EndUpdateCharacters(charProps); }
}
```

**Step 2: Implement**

```csharp
case Markdig.Extensions.Tables.Table table:
    WriteTable(ctx, table);
    break;
```

```csharp
private static void WriteTable(ConversionContext ctx, Markdig.Extensions.Tables.Table table)
{
    var doc = ctx.Document;
    var rows = table.OfType<Markdig.Extensions.Tables.TableRow>().ToList();
    if (rows.Count == 0) return;
    var colCount = rows.Max(r => r.Count);
    var dxTable = doc.Tables.Create(doc.Range.End, rows.Count, colCount);

    for (int r = 0; r < rows.Count; r++)
    {
        var mdRow = rows[r];
        for (int c = 0; c < mdRow.Count; c++)
        {
            var mdCell = (Markdig.Extensions.Tables.TableCell)mdRow[c];
            var dxCell = dxTable.Rows[r].Cells[c];

            foreach (var child in mdCell)
            {
                if (child is ParagraphBlock p && p.Inline is not null)
                    foreach (var inline in p.Inline)
                        WriteInlineIntoCell(ctx, dxCell, inline);
            }

            if (rows[r].IsHeader)
            {
                dxCell.BackgroundColor = System.Drawing.Color.FromArgb(0xF2, 0xF2, 0xF2);
                var props = doc.BeginUpdateCharacters(dxCell.Range);
                try { props.FontBold = true; }
                finally { doc.EndUpdateCharacters(props); }
            }
        }
    }
}

private static void WriteInlineIntoCell(ConversionContext ctx, TableCell dxCell, Markdig.Syntax.Inlines.Inline inline)
{
    // For now, append as a simple inline in the cell's first paragraph.
    var firstPara = dxCell.Range.Document?.Paragraphs.FirstOrDefault(p =>
        p.Range.Start.ToInt() >= dxCell.Range.Start.ToInt() &&
        p.Range.End.ToInt() <= dxCell.Range.End.ToInt())
        ?? ctx.Document.Paragraphs.Last();
    WriteInline(ctx, firstPara, inline);
}
```

The cell-paragraph lookup is fiddly — if the helper above doesn't compile cleanly with the available DevExpress API, replace `WriteInlineIntoCell` with a direct `doc.InsertText(dxCell.Range.End, ...)` for literals plus emphasis flags applied in a `BeginUpdateCharacters` block. The point of the test is structural correctness; we'll polish in Task 11 after column alignment lands.

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): GFM pipe tables with bold+shaded header"
```

---

### Task 11: GFM pipe table — column alignment

**Step 1: Test**
```csharp
[Fact]
public void Pipe_table_column_alignment_from_gfm_spec()
{
    var md = "| L | C | R |\n|:---|:---:|---:|\n| a | b | c |";
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);

    var table = server.Document.Tables[0];
    var dataRow = table.Rows[1];

    AlignmentOf(server.Document, dataRow.Cells[0]).Should().Be(ParagraphAlignment.Left);
    AlignmentOf(server.Document, dataRow.Cells[1]).Should().Be(ParagraphAlignment.Center);
    AlignmentOf(server.Document, dataRow.Cells[2]).Should().Be(ParagraphAlignment.Right);

    static ParagraphAlignment AlignmentOf(Document doc, TableCell cell)
    {
        var props = doc.BeginUpdateParagraphs(cell.Range);
        try { return props.Alignment; }
        finally { doc.EndUpdateParagraphs(props); }
    }
}
```

**Step 2: Implement** — read `table.ColumnDefinitions[c].Alignment` (Markdig type `TableColumnAlign?`), translate, apply via `BeginUpdateParagraphs(dxCell.Range).Alignment = …`.

```csharp
// In WriteTable, after writing cell content:
if (table.ColumnDefinitions is { } cols && c < cols.Count && cols[c].Alignment is { } align)
{
    var pProps = doc.BeginUpdateParagraphs(dxCell.Range);
    try
    {
        pProps.Alignment = align switch
        {
            Markdig.Extensions.Tables.TableColumnAlign.Left   => ParagraphAlignment.Left,
            Markdig.Extensions.Tables.TableColumnAlign.Center => ParagraphAlignment.Center,
            Markdig.Extensions.Tables.TableColumnAlign.Right  => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left,
        };
    }
    finally { doc.EndUpdateParagraphs(pProps); }
}
```

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): GFM table column alignment"
```

---

# Phase C — Inline mapping

> Same TDD pattern; tests added to `MarkdownToDocxConverterTests`.

### Task 12: Bold / italic / bold-italic

**Step 1: Test**
```csharp
[Fact]
public void Emphasis_produces_bold_italic_runs()
{
    var md = "**bold** *italic* ***both***";
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);

    var text = server.Document.GetText(server.Document.Range);
    text.Should().Contain("bold").And.Contain("italic").And.Contain("both");

    bool HasRun(Func<DevExpress.XtraRichEdit.API.Native.CharacterProperties, bool> predicate)
    {
        var doc = server.Document;
        for (int i = 0; i < doc.Range.End.ToInt(); i++)
        {
            var range = doc.CreateRange(doc.Range.Start.ToInt() + i, 1);
            var props = doc.BeginUpdateCharacters(range);
            try { if (predicate(props)) return true; }
            finally { doc.EndUpdateCharacters(props); }
        }
        return false;
    }
    HasRun(p => p.FontBold == true && p.FontItalic != true).Should().BeTrue();
    HasRun(p => p.FontItalic == true && p.FontBold != true).Should().BeTrue();
    HasRun(p => p.FontBold == true && p.FontItalic == true).Should().BeTrue();
}
```

**Step 2: Implement** — extend `WriteInline`:

```csharp
case Markdig.Syntax.Inlines.EmphasisInline em:
{
    var startPos = para.Range.End.ToInt();
    foreach (var child in em)
        WriteInline(ctx, para, child);
    var endPos = para.Range.End.ToInt();
    if (endPos <= startPos) break;

    var range = ctx.Document.CreateRange(startPos, endPos - startPos);
    var props = ctx.Document.BeginUpdateCharacters(range);
    try
    {
        if (em.DelimiterCount >= 2) props.FontBold = true;
        if (em.DelimiterCount == 1 || em.DelimiterCount == 3) props.FontItalic = true;
    }
    finally { ctx.Document.EndUpdateCharacters(props); }
    break;
}
```

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): bold + italic + bold-italic emphasis"
```

---

### Task 13: Inline code

**Step 1: Test**
```csharp
[Fact]
public void Inline_code_run_uses_Consolas()
{
    var md = "x `code` y";
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);

    var doc = server.Document;
    bool foundConsolasRun = false;
    var text = doc.GetText(doc.Range);
    var idx = text.IndexOf("code", StringComparison.Ordinal);
    if (idx >= 0)
    {
        var range = doc.CreateRange(doc.Range.Start.ToInt() + idx, "code".Length);
        var props = doc.BeginUpdateCharacters(range);
        try { foundConsolasRun = props.FontName == "Consolas"; }
        finally { doc.EndUpdateCharacters(props); }
    }
    foundConsolasRun.Should().BeTrue();
}
```

**Step 2: Implement** — extend `WriteInline`:

```csharp
case Markdig.Syntax.Inlines.CodeInline code:
{
    var insertedAt = ctx.Document.InsertText(para.Range.End, code.Content);
    var props = ctx.Document.BeginUpdateCharacters(insertedAt);
    try
    {
        props.FontName = "Consolas";
        props.FontSize = 9f;
        // Background colour: try props.BackColor = Color.FromArgb(0xF2,0xF2,0xF2). If
        // not exposed in this DevExpress version, leave font-only (see design Risk 1).
    }
    finally { ctx.Document.EndUpdateCharacters(props); }
    break;
}
```

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): inline code (Consolas)"
```

---

### Task 14: Hyperlinks + autolinks

**Step 1: Test**
```csharp
[Fact]
public void Hyperlink_emits_field_with_target()
{
    var md = "see [the docs](https://example.com/x) here";
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);

    server.Document.Hyperlinks.Count.Should().BeGreaterThan(0);
    server.Document.Hyperlinks[0].NavigateUri.Should().Be("https://example.com/x");
}

[Fact]
public void Autolink_emits_hyperlink_with_url_as_text()
{
    var md = "see <https://example.com/y>";
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);
    server.Document.Hyperlinks.Should().Contain(h => h.NavigateUri == "https://example.com/y");
}
```

**Step 2: Implement** — extend `WriteInline`:

```csharp
case Markdig.Syntax.Inlines.LinkInline link when !link.IsImage:
{
    var displayText = string.Concat(
        link.OfType<Markdig.Syntax.Inlines.LiteralInline>().Select(l => l.Content.ToString()));
    if (string.IsNullOrEmpty(displayText)) displayText = link.Url ?? "";
    var insertedAt = ctx.Document.InsertText(para.Range.End, displayText);
    var hl = ctx.Document.Hyperlinks.Create(insertedAt);
    hl.NavigateUri = link.Url ?? "";
    break;
}

case Markdig.Syntax.Inlines.AutolinkInline autolink:
{
    var insertedAt = ctx.Document.InsertText(para.Range.End, autolink.Url);
    var hl = ctx.Document.Hyperlinks.Create(insertedAt);
    hl.NavigateUri = autolink.Url;
    break;
}
```

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): hyperlinks + autolinks"
```

---

### Task 15: Hard line break inside paragraph

**Step 1: Test**
```csharp
[Fact]
public void Hard_break_inserts_line_break_inside_paragraph()
{
    var md = "line one  \nline two";   // two trailing spaces -> hard break
    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, null);

    var text = server.Document.GetText(server.Document.Range);
    text.Should().Contain("line one").And.Contain("line two");
    text.Should().Contain("\v");   // line break char
}
```

**Step 2: Implement** — extend `WriteInline`:

```csharp
case Markdig.Syntax.Inlines.LineBreakInline br:
    if (br.IsHard)
        ctx.Document.InsertText(para.Range.End, "\v");
    else
        ctx.Document.InsertText(para.Range.End, " ");
    break;
```

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): hard + soft line breaks"
```

---

### Task 16: Images — local embed, missing local drop, remote drop

**Step 1: Tests** (three sub-tests in one go)

```csharp
[Fact]
public void Image_local_file_is_embedded()
{
    using var tmpDir = new TempDir();
    var pngPath = Path.Combine(tmpDir.Path, "dot.png");
    File.WriteAllBytes(pngPath, OnePixelPng());
    var md = $"![dot](dot.png)";

    using var server = new RichEditDocumentServer();
    MarkdownToDocxConverter.Apply(server.Document, md, tmpDir.Path);
    server.Document.Images.Count.Should().Be(1);
}

[Fact]
public void Image_missing_local_file_is_dropped_no_throw()
{
    using var tmpDir = new TempDir();
    var md = "![](missing.png)";
    using var server = new RichEditDocumentServer();
    var act = () => MarkdownToDocxConverter.Apply(server.Document, md, tmpDir.Path);
    act.Should().NotThrow();
    server.Document.Images.Count.Should().Be(0);
}

[Fact]
public void Image_remote_url_is_dropped_no_throw()
{
    var md = "![](https://example.com/x.png)";
    using var server = new RichEditDocumentServer();
    var act = () => MarkdownToDocxConverter.Apply(server.Document, md, null);
    act.Should().NotThrow();
    server.Document.Images.Count.Should().Be(0);
}

private static byte[] OnePixelPng() => Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

private sealed class TempDir : IDisposable
{
    public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
        $"mcpoffice-md-{Guid.NewGuid():N}");
    public TempDir() => Directory.CreateDirectory(Path);
    public void Dispose() { try { Directory.Delete(Path, true); } catch { } }
}
```

**Step 2: Implement** — extend `WriteInline`:

```csharp
case Markdig.Syntax.Inlines.LinkInline imgLink when imgLink.IsImage:
{
    if (TryResolveLocalImage(imgLink.Url, ctx.BaseDirectory, out var resolved))
    {
        using var stream = File.OpenRead(resolved!);
        ctx.Document.Images.Append(stream);
    }
    // Remote / missing: silently dropped (Serilog warning added in Task 21 wire-up).
    break;
}
```

```csharp
private static bool TryResolveLocalImage(string? url, string? baseDir, out string? resolved)
{
    resolved = null;
    if (string.IsNullOrWhiteSpace(url)) return false;
    if (url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        url.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) return false;

    var candidate = Path.IsPathFullyQualified(url) ? url : Path.Combine(baseDir ?? "", url);
    if (!File.Exists(candidate)) return false;
    resolved = candidate;
    return true;
}
```

**Step 3: Run, commit**
```bash
git commit -am "feat(markdown): image embed (local) + drop (remote/missing)"
```

---

# Phase D — Wire-up + migration

### Task 17: Swap `CreateFromMarkdown` and `AppendMarkdown` to use new converter

**Files:**
- Modify: `src/mcpOffice/Services/Word/WordDocumentService.cs`

**Step 1:** Replace the bodies of `CreateFromMarkdown` and `AppendMarkdown`:

```csharp
public string CreateFromMarkdown(string path, string markdown, bool overwrite)
{
    PathGuard.RequireWritable(path, overwrite);
    try
    {
        using var server = new RichEditDocumentServer();
        var baseDir = Path.GetDirectoryName(path);
        MarkdownToDocxConverter.Apply(server.Document, markdown ?? string.Empty, baseDir);
        server.SaveDocument(path, RichEditFormat.OpenXml);
        return path;
    }
    catch (Exception ex) when (ex is not McpException)
    {
        throw ToolError.IoError(ex.Message);
    }
}

public string AppendMarkdown(string path, string markdown)
{
    PathGuard.RequireExists(path);
    try
    {
        using var server = LoadOpenXml(path);
        var baseDir = Path.GetDirectoryName(path);
        MarkdownToDocxConverter.Apply(server.Document, markdown ?? string.Empty, baseDir);
        server.SaveDocument(path, RichEditFormat.OpenXml);
        return path;
    }
    catch (Exception ex) when (ex is not McpException)
    {
        throw ToolError.IoError(ex.Message);
    }
}
```

**Step 2:** Run **only** the existing markdown tests first to surface any regressions:
```bash
dotnet test tests/mcpOffice.Tests --filter "FullyQualifiedName~CreateFromMarkdownTests|FullyQualifiedName~AppendMarkdownTests|FullyQualifiedName~MarkdownTests" --nologo
```
If anything fails, inspect each — adjust the assertion if it was testing an artifact of the old converter (e.g. exact paragraph count including trailing blank); fix the converter if it's a real correctness gap.

**Step 3: Run full test suite**
```bash
dotnet test --nologo
```
Expected: all green.

**Step 4: Commit**
```bash
git commit -am "feat: rewire CreateFromMarkdown + AppendMarkdown to Markdig"
```

---

### Task 18: Swap markdown branch of `Convert`

**Files:**
- Modify: `src/mcpOffice/Services/Word/WordDocumentService.cs` (the `Convert` method's `WordOutputFormat.Markdown`-as-input handling)

**Step 1:** Locate where `Convert` reads a `.md` input. Currently markdown is only an *output* format; the input side comes via `srv.LoadDocument(...)` which uses the DevExpress format inferred from extension. For `.md` input, we can't trust DevExpress — call our converter instead.

Replace the input-side load (around the `Convert` method head) with a branch:
```csharp
if (Path.GetExtension(inputPath).Equals(".md", StringComparison.OrdinalIgnoreCase) ||
    Path.GetExtension(inputPath).Equals(".markdown", StringComparison.OrdinalIgnoreCase))
{
    using var server = new RichEditDocumentServer();
    var md = File.ReadAllText(inputPath, Encoding.UTF8);
    MarkdownToDocxConverter.Apply(server.Document, md, Path.GetDirectoryName(inputPath));
    // ... existing output-format save block reused, taking `server` as the populated document
}
else
{
    // existing path
}
```

The exact diff depends on the current shape of `Convert` — read lines 309-370 of `WordDocumentService.cs` first and refactor to share the output-save block between both input branches. Keep the output-format dispatch (`switch (resolved)`) untouched.

**Step 2: Test** — add or extend a `Convert` test that uses an `.md` input file with inline code + a table, asserts the resulting `.docx` has `Tables.Count >= 1` and the inline-code text survives.

**Step 3: Run full suite, commit**
```bash
dotnet test --nologo
git commit -am "feat: word_convert .md input uses Markdig path"
```

---

### Task 19: Remove `MarkdownToDocxGenerator` package + dead helpers

**Files:**
- Modify: `src/mcpOffice/mcpOffice.csproj`
- Modify: `src/mcpOffice/Services/Word/WordDocumentService.cs` — delete `CreateDocumentFromMarkdown`, `NormalizeMarkdownGeneratedDocument`, `ApplyMarkdownHeadingStyles`, `ApplyMarkdownItalicStyles`, `ExtractMarkdownItalicSpans`, the `MarkdownHeading` record, and the `using MarkdownToDocxGenerator;` directive.

**Step 1: Remove the package**
```bash
dotnet remove src/mcpOffice package MarkdownToDocxGenerator
```

**Step 2: Delete dead code in `WordDocumentService.cs`** — grep for each helper name and delete its definition, then delete the `using` and the `MarkdownHeading` record. Keep `EnsureParagraphStyle` in `WordDocumentService` (still used by other tools); the converter has its own copy — de-dupe by making the converter call the service's helper, OR move the helper to a shared internal static class. Pick whichever is one-line cheaper and document the choice in the commit message.

**Step 3: Build clean + run full tests**
```bash
dotnet build --nologo
dotnet test --nologo
```
Expected: 0 warnings, all green.

**Step 4: Commit**
```bash
git commit -am "chore: remove MarkdownToDocxGenerator + post-process helpers"
```

---

# Phase E — Real-world fidelity + final verification

### Task 20: Real-world fidelity test (`fn_send_email_callers.md`)

**Files:**
- Create: `tests/fixtures/fn_send_email_callers.md` — copy of `C:\Projects\LimsBasic\docs\fn_send_email_callers.md`
- Create: `tests/mcpOffice.Tests/Word/MarkdownRealWorldTests.cs`

**Step 1: Copy fixture**
```bash
cp "C:\Projects\LimsBasic\docs\fn_send_email_callers.md" tests/fixtures/fn_send_email_callers.md
```

**Step 2: Test**
```csharp
// tests/mcpOffice.Tests/Word/MarkdownRealWorldTests.cs
using DevExpress.XtraRichEdit;
using FluentAssertions;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class MarkdownRealWorldTests
{
    [Fact]
    public void Fn_send_email_callers_md_round_trips_with_tables_and_inline_code()
    {
        var md = File.ReadAllText(TestFixtures.Path("fn_send_email_callers.md"));
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        // Tables present
        server.Document.Tables.Count.Should().BeGreaterThanOrEqualTo(4);

        // Inline code preserved (FN_SEND_EMAIL appears outside any heading)
        var bodyText = server.Document.GetText(server.Document.Range);
        bodyText.Should().Contain("FN_SEND_EMAIL");

        // Bold survived (**optional**)
        bool anyBoldRun = false;
        var doc = server.Document;
        var totalChars = doc.Range.End.ToInt() - doc.Range.Start.ToInt();
        // Sample every 10 chars — exhaustive scan is overkill
        for (int i = 0; i < totalChars && !anyBoldRun; i += 10)
        {
            var range = doc.CreateRange(doc.Range.Start.ToInt() + i, 1);
            var props = doc.BeginUpdateCharacters(range);
            try { if (props.FontBold == true) anyBoldRun = true; }
            finally { doc.EndUpdateCharacters(props); }
        }
        anyBoldRun.Should().BeTrue();
    }
}
```

**Step 3: Run**
```bash
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~MarkdownRealWorldTests --nologo
```
Expected: pass.

**Step 4: Commit**
```bash
git add tests/fixtures/fn_send_email_callers.md tests/mcpOffice.Tests/Word/MarkdownRealWorldTests.cs
git commit -m "test: real-world fidelity for Markdig path (LimsBasic fn_send_email_callers)"
```

---

### Task 21: Final verification + Serilog warning + handoff

**Step 1: Add the Serilog warning hooks** previewed in the design (only one is observable: image dropped). Inject an `ILogger`-like callback into `MarkdownToDocxConverter.Apply` only if convenient — otherwise skip (these are debug aids, not user-facing contract). Decide based on whether DevExpress test setup makes logger injection painful.

**Step 2: Live MCP smoke** — re-convert `C:\Projects\LimsBasic\docs\fn_send_email_callers.md` via the running server (Claude Code `mcp__office__word_convert` or stdio), open the resulting `.docx` in Word, confirm visually:
- Inline code spans render in Consolas
- The four tables exist with bordered cells and bold/shaded headers
- `**optional**` is bold

**Step 3: Final build+test**
```bash
dotnet build -c Release --nologo
dotnet test -c Release --nologo
```
Expected: 0 warnings, all green.

**Step 4: Update `SESSION_HANDOFF.md` and `TODO.md`** — note the converter swap, the four risk items resolved (or carried over if any DevExpress API fell back), the new test count, and clear the v2 entry from TODO.

**Step 5: PR**
```bash
git push -u origin feat/markdown-to-docx-markdig
gh pr create --title "feat: Markdig-based markdown -> docx converter" --body "$(cat <<'EOF'
## Summary

- Replaces `MarkdownToDocxGenerator` v1.2.0 with a Markdig parse + custom AST walker that emits DevExpress `Document` API calls directly.
- Fixes inline-code stripping, GFM table flattening, and bold-run loss observed in real-world docs.
- Removes ~150 lines of regex post-process patches.

## Test plan

- [x] All 16 new converter unit tests pass
- [x] Existing CreateFromMarkdown / AppendMarkdown / Markdown tests pass unchanged
- [x] Real-world fixture (`fn_send_email_callers.md`) round-trips with tables + inline code + bold preserved
- [x] Full suite (unit + integration) green
- [ ] Live MCP smoke: convert `fn_send_email_callers.md` via stdio, open in Word, eyeball

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

---

## What this plan deliberately does NOT do

- No syntax highlighting in fenced code blocks.
- No support for footnotes, task lists, strikethrough, math, or definition lists.
- No fix to the read-back path (`word_read_markdown`) — separate fix.
- No HTML-in-markdown parsing — literal text only.
- No remote image fetching.
- No de-duplication of `EnsureParagraphStyle` beyond what Task 19 calls out — over-engineering for a POC.

## Risks called out (carried from design doc)

1. **DevExpress character shading on inline code may not be exposed.** Fallback in Task 13: keep Consolas font, drop the gray background, leave a code comment.
2. **GFM column alignment** — Task 11 verifies this works; if `BeginUpdateParagraphs(cell.Range).Alignment` doesn't propagate, fall back to per-cell paragraph alignment loops.
3. **Existing test assertions may be tuned to the old converter's idiosyncrasies.** Task 17 surfaces these by running the markdown test subset before the full suite.
