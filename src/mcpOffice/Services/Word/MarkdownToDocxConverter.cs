using System.Drawing;
using DevExpress.Office.Utils;
using DevExpress.XtraRichEdit.API.Native;
using Markdig;
using Markdig.Extensions.Tables;
using MdTable = Markdig.Extensions.Tables.Table;
using MdTableRow = Markdig.Extensions.Tables.TableRow;
using MdTableCell = Markdig.Extensions.Tables.TableCell;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace McpOffice.Services.Word;

internal static class MarkdownToDocxConverter
{
    private static readonly MarkdownPipeline Pipeline =
        new MarkdownPipelineBuilder().UsePipeTables().Build();

    public static void Apply(Document document, string markdown, string? baseDirectory, bool preserveExistingHeadingStyles = false)
    {
        var ast = Markdown.Parse(markdown, Pipeline);
        var ctx = new ConversionContext(document, baseDirectory)
        {
            PreserveExistingHeadingStyles = preserveExistingHeadingStyles,
        };
        foreach (var block in ast)
            WriteBlock(ctx, block);
    }

    private sealed class ConversionContext(Document Document, string? BaseDirectory)
    {
        public Document Document { get; } = Document;
        public string? BaseDirectory { get; } = BaseDirectory;

        // When the document was seeded from a .dotx/.docx template, heading styles the
        // template defines must win over the converter's built-in heading formatting.
        public bool PreserveExistingHeadingStyles { get; init; }

        // Accumulated emphasis depth from enclosing EmphasisInline nodes.
        // Bold when boldDepth > 0; Italic when italicDepth > 0.
        public int BoldDepth { get; set; }
        public int ItalicDepth { get; set; }
    }

    private static void WriteBlock(ConversionContext ctx, Block block)
    {
        switch (block)
        {
            case HeadingBlock h:
                WriteHeading(ctx, h);
                break;
            case ParagraphBlock p:
                WriteParagraph(ctx, p);
                break;
            case ListBlock list:
                WriteList(ctx, list, level: 0);
                break;
            case QuoteBlock q:
                WriteQuote(ctx, q);
                break;
            case ThematicBreakBlock:
                WriteHorizontalRule(ctx);
                break;
            case FencedCodeBlock fenced:
                WriteCodeBlock(ctx, fenced.Lines.ToString());
                break;
            case CodeBlock code:
                WriteCodeBlock(ctx, code.Lines.ToString());
                break;
            case MdTable mdTable:
                WriteTable(ctx, mdTable);
                break;
            // Other block kinds added in subsequent tasks.
            default:
                // Unknown blocks silently skipped; Serilog warning attached in Task 21.
                break;
        }
    }

    private static void WriteHeading(ConversionContext ctx, HeadingBlock block)
    {
        var level = Math.Clamp(block.Level, 1, 6);
        var styleName = $"Heading {level}";
        EnsureHeadingStyle(ctx.Document, level, styleName, ctx.PreserveExistingHeadingStyles);
        var para = AppendNewParagraph(ctx);
        para.Style = ctx.Document.ParagraphStyles[styleName];
        if (block.Inline is null) return;
        foreach (var inline in block.Inline)
            WriteInline(ctx, para, inline);
    }

    // Word's classic heading palette — dark blue H1, medium blue H2-H6.
    private static readonly Color HeadingDarkBlue = Color.FromArgb(0x1F, 0x38, 0x64);
    private static readonly Color HeadingBlue     = Color.FromArgb(0x2E, 0x74, 0xB5);

    private static void EnsureHeadingStyle(Document doc, int level, string styleName, bool preserveExisting)
    {
        // A fresh RichEditDocumentServer ships without populated built-in heading styles
        // (per DevExpress docs). Whether or not a style exists, always (re)apply formatting
        // so the rendered output has real heading hierarchy instead of body-text-with-a-name.
        var style = doc.ParagraphStyles[styleName];
        if (style is not null && preserveExisting)
        {
            // Style comes from a user-supplied template — its formatting is authoritative.
            return;
        }
        if (style is null)
        {
            style = doc.ParagraphStyles.CreateNew();
            style.Name = styleName;
            doc.ParagraphStyles.Add(style);
        }
        // DevExpress OutlineLevel is 1-based: 0 = no outline (body text),
        // 1 = Heading 1, ..., 9 = Heading 9. OOXML serializes as level-1
        // (Heading 1 = outlineLvl 0). Set the DevExpress value to match the
        // markdown heading depth directly.
        style.OutlineLevel = level;
        switch (level)
        {
            case 1:
                style.FontSize = 16f;
                style.Bold = true;
                style.ForeColor = HeadingDarkBlue;
                style.SpacingBefore = Units.InchesToDocumentsF(0.17f); // ~12pt
                style.SpacingAfter  = Units.InchesToDocumentsF(0.06f); // ~4pt
                break;
            case 2:
                style.FontSize = 13f;
                style.Bold = true;
                style.ForeColor = HeadingBlue;
                style.SpacingBefore = Units.InchesToDocumentsF(0.14f); // ~10pt
                style.SpacingAfter  = Units.InchesToDocumentsF(0.06f);
                break;
            case 3:
                style.FontSize = 12f;
                style.Bold = true;
                style.ForeColor = HeadingBlue;
                style.SpacingBefore = Units.InchesToDocumentsF(0.11f); // ~8pt
                style.SpacingAfter  = Units.InchesToDocumentsF(0.04f);
                break;
            case 4:
                style.FontSize = 11f;
                style.Bold = true;
                style.ForeColor = HeadingBlue;
                style.SpacingBefore = Units.InchesToDocumentsF(0.08f);
                style.SpacingAfter  = Units.InchesToDocumentsF(0.03f);
                break;
            case 5:
                style.FontSize = 11f;
                style.Italic = true;
                style.ForeColor = HeadingBlue;
                break;
            case 6:
                style.FontSize = 11f;
                style.Italic = true;
                style.ForeColor = HeadingDarkBlue;
                break;
        }
    }

    private static void EnsureParagraphStyle(Document doc, string styleName)
    {
        if (doc.ParagraphStyles[styleName] is not null) return;
        var s = doc.ParagraphStyles.CreateNew();
        s.Name = styleName;
        doc.ParagraphStyles.Add(s);
    }

    private static void WriteParagraph(ConversionContext ctx, ParagraphBlock block)
    {
        var para = AppendNewParagraph(ctx);
        if (block.Inline is null) return;
        foreach (var inline in block.Inline)
            WriteInline(ctx, para, inline);
    }

    private static void WriteQuote(ConversionContext ctx, QuoteBlock block)
    {
        foreach (var child in block)
        {
            if (child is ParagraphBlock p)
            {
                var para = AppendNewParagraph(ctx);
                // 0.25" expressed in DevExpress document units (1/300th of an inch).
                para.LeftIndent = Units.InchesToDocumentsF(0.25f);
                if (p.Inline is null) continue;
                foreach (var inline in p.Inline)
                    WriteInline(ctx, para, inline);
            }
        }
    }

    private static Paragraph AppendNewParagraph(ConversionContext ctx)
    {
        var doc = ctx.Document;
        // DevExpress Document has no InsertParagraph(DocumentPosition); follow the existing
        // project pattern (WordDocumentService.InsertParagraph) of inserting "\n".
        doc.InsertText(doc.Range.End, "\n");
        var para = doc.Paragraphs[doc.Paragraphs.Count - 1];
        // The newly-inserted paragraph inherits the previous paragraph's style AND list state;
        // reset both so each writer starts with a blank slate. WriteHeading and WriteList
        // immediately re-set style / list properties; everything else (paragraphs, code blocks,
        // blockquotes, hrs) gets a clean Normal-styled non-list paragraph.
        var normalStyle = doc.ParagraphStyles["Normal"] ?? doc.ParagraphStyles["Default Paragraph Style"];
        if (normalStyle is not null) para.Style = normalStyle;
        para.ListIndex = -1;
        para.ListLevel = 0;
        // Direct paragraph formatting (borders from a horizontal rule, indent from a
        // quote/code block) is inherited by the new paragraph as well; clear it so the
        // border of one `---` doesn't underline every paragraph after it. Writers that
        // want a border/indent (hr, code block, quote) set it after this reset.
        ClearDirectParagraphFormatting(doc, para.Range);
        return para;
    }

    private static void ClearDirectParagraphFormatting(Document doc, DocumentRange range)
    {
        var props = doc.BeginUpdateParagraphs(range);
        try
        {
            props.LeftIndent = 0f;
            props.Borders.TopBorder.LineStyle = BorderLineStyle.None;
            props.Borders.BottomBorder.LineStyle = BorderLineStyle.None;
            props.Borders.LeftBorder.LineStyle = BorderLineStyle.None;
            props.Borders.RightBorder.LineStyle = BorderLineStyle.None;
        }
        finally { doc.EndUpdateParagraphs(props); }
    }

    private static void WriteList(ConversionContext ctx, ListBlock list, int level)
    {
        var doc = ctx.Document;

        // Create the abstract numbering list from the appropriate template.
        var template = list.IsOrdered
            ? doc.AbstractNumberingLists.NumberedListTemplate
            : doc.AbstractNumberingLists.BulletedListTemplate;
        var abstractList = template.CreateNew();
        doc.AbstractNumberingLists.Add(abstractList);

        // Create the concrete numbering list that references the abstract one.
        var numberingList = doc.NumberingLists.Add(abstractList.Index);
        var listIndex = numberingList.Index;

        foreach (var item in list.OfType<ListItemBlock>())
        {
            foreach (var sub in item)
            {
                switch (sub)
                {
                    case ParagraphBlock p:
                    {
                        var para = AppendNewParagraph(ctx);
                        para.ListIndex = listIndex;
                        para.ListLevel = level;
                        if (p.Inline is null) break;
                        foreach (var inline in p.Inline)
                            WriteInline(ctx, para, inline);
                        break;
                    }
                    case ListBlock nested:
                        WriteList(ctx, nested, level + 1);
                        break;
                }
            }
        }
    }

    private static readonly Color CodeBackground = Color.FromArgb(0xF2, 0xF2, 0xF2);

    private static readonly Color CodeBorder = Color.FromArgb(0xC0, 0xC0, 0xC0);

    private static void WriteCodeBlock(ConversionContext ctx, string text)
    {
        var doc = ctx.Document;
        var lines = text.Replace("\r\n", "\n").Split('\n');
        foreach (var line in lines)
        {
            var para = AppendNewParagraph(ctx);
            para.LeftIndent = Units.InchesToDocumentsF(0.15f);

            // Visual marker: a left border that runs alongside the indent.
            // Without this the indent reads like a quote rather than a code block.
            var pProps = doc.BeginUpdateParagraphs(para.Range);
            try
            {
                pProps.Borders.LeftBorder.LineStyle = BorderLineStyle.Single;
                pProps.Borders.LeftBorder.LineWidth = 1.5f;
                pProps.Borders.LeftBorder.LineColor = CodeBorder;
            }
            finally { doc.EndUpdateParagraphs(pProps); }

            if (line.Length == 0) continue;

            var insertedRange = doc.InsertText(para.Range.End, line);
            var props = doc.BeginUpdateCharacters(insertedRange);
            try
            {
                props.FontName = "Consolas";
                props.FontSize = 9f;
                props.BackColor = CodeBackground;
            }
            finally { doc.EndUpdateCharacters(props); }
        }
    }

    private static void WriteHorizontalRule(ConversionContext ctx)
    {
        var para = AppendNewParagraph(ctx);
        var props = ctx.Document.BeginUpdateParagraphs(para.Range);
        try
        {
            props.Borders.BottomBorder.LineStyle = BorderLineStyle.Single;
            props.Borders.BottomBorder.LineWidth = 0.5f;
        }
        finally { ctx.Document.EndUpdateParagraphs(props); }
    }

    private static readonly Color HeaderBackground = Color.FromArgb(0xF2, 0xF2, 0xF2);

    private static void WriteTable(ConversionContext ctx, MdTable table)
    {
        var doc = ctx.Document;
        var rows = table.OfType<MdTableRow>().ToList();
        if (rows.Count == 0) return;

        var colCount = rows.Max(r => r.Count);
        if (colCount == 0) return;

        var dxTable = doc.Tables.Create(doc.Range.End, rows.Count, colCount);

        // Freshly created cell paragraphs inherit the style and direct formatting of
        // the paragraph preceding the table (often a heading, or a border left over
        // from a horizontal rule), leaking heading formatting, outline levels and
        // bottom borders into every cell. Reset to Normal, mirroring AppendNewParagraph.
        var normalStyle = doc.ParagraphStyles["Normal"] ?? doc.ParagraphStyles["Default Paragraph Style"];
        if (normalStyle is not null)
        {
            foreach (var para in doc.Paragraphs.Get(dxTable.Range))
                para.Style = normalStyle;
        }
        ClearDirectParagraphFormatting(doc, dxTable.Range);

        for (int r = 0; r < rows.Count; r++)
        {
            var mdRow = rows[r];
            for (int c = 0; c < mdRow.Count; c++)
            {
                var mdCell = (MdTableCell)mdRow[c];
                var dxCell = dxTable.Rows[r].Cells[c];

                var isHeader = mdRow.IsHeader;
                if (isHeader)
                    dxCell.BackgroundColor = HeaderBackground;

                // Use a tracked cursor position anchored to the cell's live ContentRange.Start.
                // We re-read from dxCell each time because earlier-cell insertions shift absolute
                // positions — but the dxTable/dxCell object always returns the current live position.
                // After each WriteInline call, cursor advances to the end of the last insertion
                // so subsequent inlines append correctly (same cell, not beginning).
                if (isHeader) ctx.BoldDepth++;
                var cursor = new CellCursor(dxCell);
                foreach (var child in mdCell)
                {
                    if (child is ParagraphBlock p && p.Inline is not null)
                    {
                        foreach (var inline in p.Inline)
                            WriteCellInline(ctx, cursor, inline);
                    }
                }
                if (isHeader) ctx.BoldDepth--;

                // Apply GFM column alignment (`:---` left, `:---:` center, `---:` right).
                if (table.ColumnDefinitions is { } cols && c < cols.Count && cols[c].Alignment is { } align)
                {
                    var pProps = doc.BeginUpdateParagraphs(dxCell.ContentRange);
                    try
                    {
                        pProps.Alignment = align switch
                        {
                            TableColumnAlign.Left   => ParagraphAlignment.Left,
                            TableColumnAlign.Center => ParagraphAlignment.Center,
                            TableColumnAlign.Right  => ParagraphAlignment.Right,
                            _                       => ParagraphAlignment.Left,
                        };
                    }
                    finally { doc.EndUpdateParagraphs(pProps); }
                }
            }
        }
    }

    /// <summary>
    /// Tracks the insertion cursor within a table cell.
    /// Re-reads <see cref="TableCell.ContentRange.Start"/> from the live DevExpress cell reference
    /// for the first character; subsequent insertions advance the cursor to the end of the
    /// previous inserted range so inlines append rather than prepend.
    /// </summary>
    private sealed class CellCursor(DevExpress.XtraRichEdit.API.Native.TableCell cell)
    {
        private DocumentPosition? _position;

        public DocumentPosition Current =>
            _position ?? cell.ContentRange.Start;

        public void Advance(DocumentRange inserted) =>
            _position = inserted.End;
    }

    // RichEdit keeps a font name per script slot (ASCII, high-ANSI, complex script, East Asian).
    // Setting FontName writes all of them; resetting only CharacterPropertiesMask.FontName leaves
    // the slots set, which is why the e6db964 fix never took (MD-003). Reset every slot.
    private const CharacterPropertiesMask InheritedRunMask =
        CharacterPropertiesMask.FontName
        | CharacterPropertiesMask.FontNameAscii
        | CharacterPropertiesMask.FontNameHighAnsi
        | CharacterPropertiesMask.FontNameComplexScript
        | CharacterPropertiesMask.FontNameEastAsia
        | CharacterPropertiesMask.FontSize
        | CharacterPropertiesMask.BackColor;

    /// <summary>
    /// The single append point for inline text. RichEdit extends the run to the left on insert,
    /// so every new run inherits that run's direct formatting (Consolas after a code span, etc.).
    /// Clear it here and apply only what this run should have: the emphasis context, plus the
    /// monospace face for <paramref name="code"/> runs. Every inline that inserts text must go
    /// through this method — see CLAUDE.md "RichEdit inherits everything".
    /// </summary>
    private static DocumentRange InsertRun(ConversionContext ctx, DocumentPosition at, string text, bool code = false)
    {
        var doc = ctx.Document;
        var inserted = doc.InsertText(at, text);
        var props = doc.BeginUpdateCharacters(inserted);
        try
        {
            if (code)
            {
                props.Reset(CharacterPropertiesMask.BackColor);
                props.FontName = "Consolas";
                props.FontSize = 9f;
            }
            else
            {
                props.Reset(InheritedRunMask);
            }
            props.Bold   = ctx.BoldDepth   > 0;
            props.Italic = ctx.ItalicDepth > 0;
        }
        finally { doc.EndUpdateCharacters(props); }
        return inserted;
    }

    /// <summary>
    /// Variant of <see cref="WriteInline"/> for table cells that uses a tracked
    /// <see cref="CellCursor"/> instead of a paragraph's <c>Range.End</c>.
    /// All inline types are handled identically; only the insertion anchor differs.
    /// </summary>
    private static void WriteCellInline(ConversionContext ctx, CellCursor cursor, Inline inline)
    {
        var doc = ctx.Document;
        switch (inline)
        {
            case LiteralInline lit:
            {
                var text = lit.Content.ToString();
                if (text.Length == 0) break;
                cursor.Advance(InsertRun(ctx, cursor.Current, text));
                break;
            }
            case EmphasisInline em:
            {
                if (em.DelimiterCount >= 2) ctx.BoldDepth++;
                if (em.DelimiterCount == 1 || em.DelimiterCount == 3) ctx.ItalicDepth++;
                foreach (var child in em)
                    WriteCellInline(ctx, cursor, child);
                if (em.DelimiterCount >= 2) ctx.BoldDepth--;
                if (em.DelimiterCount == 1 || em.DelimiterCount == 3) ctx.ItalicDepth--;
                break;
            }
            case CodeInline code:
            {
                cursor.Advance(InsertRun(ctx, cursor.Current, code.Content, code: true));
                break;
            }
            case LinkInline link when !link.IsImage:
            {
                var displayText = string.Concat(
                    link.Descendants<LiteralInline>().Select(l => l.Content.ToString()));
                if (string.IsNullOrEmpty(displayText)) displayText = link.Url ?? string.Empty;
                if (displayText.Length == 0) break;
                var inserted = InsertRun(ctx, cursor.Current, displayText);
                cursor.Advance(inserted);
                var hl = doc.Hyperlinks.Create(inserted);
                hl.NavigateUri = link.Url ?? string.Empty;
                break;
            }
            case AutolinkInline autolink:
            {
                var url = autolink.Url ?? string.Empty;
                if (url.Length == 0) break;
                var inserted = InsertRun(ctx, cursor.Current, url);
                cursor.Advance(inserted);
                var hl = doc.Hyperlinks.Create(inserted);
                hl.NavigateUri = url;
                break;
            }
            case LineBreakInline br:
                // Hard break: line-break-within-paragraph (\v); soft break: space.
                cursor.Advance(InsertRun(ctx, cursor.Current, br.IsHard ? "\v" : " "));
                break;
            // Images inside table cells are silently dropped (uncommon; complex to size).
        }
    }

    private static void WriteInline(ConversionContext ctx, Paragraph para, Inline inline)
    {
        switch (inline)
        {
            case LiteralInline lit:
            {
                var text = lit.Content.ToString();
                if (text.Length == 0) break;
                InsertRun(ctx, para.Range.End, text);
                break;
            }
            case EmphasisInline em:
            {
                // Push emphasis state before writing children; pop afterwards.
                // Markdig 1.x represents ***both*** as nested em(1) { em(2) { ... } }.
                if (em.DelimiterCount >= 2) ctx.BoldDepth++;
                if (em.DelimiterCount == 1 || em.DelimiterCount == 3) ctx.ItalicDepth++;
                foreach (var child in em)
                    WriteInline(ctx, para, child);
                if (em.DelimiterCount >= 2) ctx.BoldDepth--;
                if (em.DelimiterCount == 1 || em.DelimiterCount == 3) ctx.ItalicDepth--;
                break;
            }
            case CodeInline code:
            {
                InsertRun(ctx, para.Range.End, code.Content, code: true);
                break;
            }
            case LinkInline imgLink when imgLink.IsImage:
            {
                if (TryResolveLocalImage(imgLink.Url, ctx.BaseDirectory, out var resolved))
                {
                    var imgSource = DocumentImageSource.FromFile(resolved!);
                    ctx.Document.Images.Append(imgSource);
                }
                // Remote URLs and missing local files: silently dropped.
                // Serilog warning hook is added in Task 21.
                break;
            }
            case LinkInline link when !link.IsImage:
            {
                // Concatenate inner literal text as display text. Falls back to URL if empty.
                var displayText = string.Concat(
                    link.Descendants<LiteralInline>().Select(l => l.Content.ToString()));
                if (string.IsNullOrEmpty(displayText)) displayText = link.Url ?? string.Empty;
                if (displayText.Length == 0) break;

                var insertedRange = InsertRun(ctx, para.Range.End, displayText);
                var hl = ctx.Document.Hyperlinks.Create(insertedRange);
                hl.NavigateUri = link.Url ?? string.Empty;
                break;
            }
            case AutolinkInline autolink:
            {
                var url = autolink.Url ?? string.Empty;
                if (url.Length == 0) break;
                var insertedRange = InsertRun(ctx, para.Range.End, url);
                var hl = ctx.Document.Hyperlinks.Create(insertedRange);
                hl.NavigateUri = url;
                break;
            }
            case LineBreakInline br:
                // Hard break (two trailing spaces + newline): insert \v (line-break-within-paragraph).
                // Soft break (single newline): insert a single space.
                InsertRun(ctx, para.Range.End, br.IsHard ? "\v" : " ");
                break;
        }
    }

    private static bool TryResolveLocalImage(string? url, string? baseDir, out string? resolved)
    {
        resolved = null;
        if (string.IsNullOrWhiteSpace(url)) return false;
        if (url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            url.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) return false;

        var candidate = Path.IsPathFullyQualified(url) ? url : Path.Combine(baseDir ?? string.Empty, url);
        if (!File.Exists(candidate)) return false;
        resolved = candidate;
        return true;
    }
}
