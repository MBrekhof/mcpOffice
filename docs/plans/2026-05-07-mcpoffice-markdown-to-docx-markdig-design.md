# mcpOffice — Markdown→DOCX via Markdig (Word converter v2)

**Date:** 2026-05-07
**Status:** Approved (brainstorming phase)
**Scope:** Replace the lossy `MarkdownToDocxGenerator` v1.2.0 NuGet path with a Markdig-based AST walker. Affects `word_create_from_markdown`, `word_append_markdown`, and the markdown branch of `word_convert`. Read-back path (`word_read_markdown`) is out of scope.

## Motivation

The current write-side converter (`MarkdownToDocxGenerator` v1.2.0 + post-process regex patches in `NormalizeMarkdownGeneratedDocument`) silently drops:

- Inline code spans — single backticks vanish, leaving empty placeholders in the rendered document.
- GFM pipe tables — flattened into a sequence of plain paragraphs with no row/column structure.
- Bold runs in body text — the post-process patch only restores headings + italic spans.
- Numbered/unordered list markers in some cases — gets duplicated as inline tokens.

Verified against `C:\Projects\LimsBasic\docs\fn_send_email_callers.md`: rendering of `` `FN_SEND_EMAIL` `` produced empty whitespace, and four GFM tables became orphaned paragraph runs. Headings + plain paragraphs survived.

The fix swaps the third-party converter for a Markdig parse + a custom AST walker that emits DevExpress `Document` API calls directly.

## Architecture

**Packages**

- Add: `Markdig` (NuGet, BSD-2-Clause, no transitive dependencies).
- Remove: `MarkdownToDocxGenerator` v1.2.0.

**New file:** `src/mcpOffice/Services/Word/MarkdownToDocxConverter.cs` — a stateless converter:

```csharp
public static class MarkdownToDocxConverter
{
    public static void Apply(Document document, string markdown, string? baseDirectory);
}
```

**Touched files**

- `WordDocumentService.cs` — `CreateFromMarkdown`, `AppendMarkdown`, the markdown branch of `Convert`. Each becomes a thin wrapper that calls `MarkdownToDocxConverter.Apply` against a fresh or loaded `Document`. The post-process helpers (`NormalizeMarkdownGeneratedDocument`, `ApplyMarkdownHeadingStyles`, `ApplyMarkdownItalicStyles`, `ExtractMarkdownItalicSpans`, `CreateDocumentFromMarkdown`) are deleted — Markdig hands us a structured AST so regex-patching the output is no longer needed.
- `mcpOffice.csproj` — package swap.

**Pipeline**

```
markdown string
  -> Markdig.Markdown.Parse(text, pipeline)
     pipeline = new MarkdownPipelineBuilder().UsePipeTables().Build()
  -> MarkdownDocument AST
  -> MarkdownToDocxConverter.Apply(document, ast, baseDir)
  -> DevExpress Document populated via Paragraphs / Tables / Images / runs
```

Stateless. Each call parses fresh. No caching, no shared mutable state.

## AST → docx mapping

### Block-level

| Markdig block | Output |
|---|---|
| `HeadingBlock` (level 1–6) | One paragraph, style `Heading {N}` (resolved via existing `EnsureHeadingStyle` helper, which already handles the "Titre N" alias DevExpress sometimes returns). Inlines walked as runs. |
| `ParagraphBlock` | One paragraph. Inlines walked as runs. |
| `ListBlock` (ordered/unordered) | One paragraph per `ListItemBlock`, bullet or number applied via DevExpress numbering list. Nested lists indent one level per depth. |
| `QuoteBlock` | Paragraphs with left indent 0.25". No left border (avoid template-style dependency). |
| `FencedCodeBlock` / `CodeBlock` | One paragraph per source line. Font `Consolas` 9pt, paragraph shading `#F2F2F2`, left indent 0.1". Language hint ignored. |
| `Table` (GFM pipe table) | DevExpress `Tables.Create(rows, cols)`. Borders: thin black on every cell. Header row: cell shading `#F2F2F2`, all runs bold. Cell alignment from GFM column spec (`:---` left, `:---:` center, `---:` right). |
| `ThematicBreakBlock` (`---`) | Empty paragraph with bottom border. |
| `HtmlBlock` | Treated as plain text paragraph (rare in agent-authored markdown; YAGNI on full HTML). |

### Inline

Walked as `Run`s appended to the active paragraph.

| Markdig inline | Output |
|---|---|
| `LiteralInline` | Plain run. |
| `EmphasisInline` (depth 1) | Run with `Italic = true`. |
| `EmphasisInline` (depth 2) | Run with `Bold = true`. |
| `EmphasisInline` (depth 3) | Run with `Bold = Italic = true`. |
| `CodeInline` | Run with `FontName = "Consolas"`, `FontSize = 9pt`, character shading `#F2F2F2` if DevExpress allows (else font alone — see Risks). |
| `LinkInline` (image=false) | Hyperlink field with display text from inner inlines. |
| `LinkInline` (image=true) | Resolve `Url` against `baseDir` if relative. If local file exists, `doc.Images.Append(stream, …)`. If http(s) or missing, log warning to stderr, drop. |
| `AutolinkInline` | Hyperlink, URL as both target and display text. |
| `LineBreakInline` | Soft break = space; hard break (`  \n`) = line-break inside the same paragraph (`\v`). |
| Anything else | Falls through as literal text. |

### Edge behavior

- Empty markdown → empty document, no exception.
- Markdig is lenient: anything unparseable becomes a `LiteralInline`. No error path here.
- Unknown block type → skip with a stderr warning.

## Tool surface

No public API change.

```
word_create_from_markdown(path, markdown, overwrite=false)   -> path
word_append_markdown(path, markdown)                         -> path
word_convert(inputPath, outputPath, format?)                 -> path   // unchanged for non-md routes
```

### Image base directory resolution

No new parameter. The base directory is inferred per tool:

| Tool | `baseDir` passed to converter |
|---|---|
| `word_create_from_markdown(path, markdown, …)` | `Path.GetDirectoryName(path)` — output's parent dir. |
| `word_append_markdown(path, markdown)` | `Path.GetDirectoryName(path)` — same logic. |
| `word_convert(input.md, output.docx, …)` | `Path.GetDirectoryName(input)` — input's parent dir. The intuitive choice when converting an existing `.md` with sibling `.png`s. |

Absolute image paths bypass `baseDir`. Remote `http(s)` URLs are dropped with a warning regardless.

### Logging

Serilog writes warnings to stderr (already configured in `Program.cs`):

- `image dropped: {url} (remote URL fetching disabled)`
- `image dropped: {path} (file not found, resolved against {baseDir})`
- `unknown block type: {type} (skipped)`

No new `McpException` codes. Bad markdown doesn't fail.

## Testing

### Existing tests

`tests/mcpOffice.Tests/Word/CreateFromMarkdownTests.cs`, `AppendMarkdownTests.cs`, `MarkdownTests.cs` (~120 lines combined) get re-run on the new converter. Most assert round-trips and outline structure, so they should pass with no edits. Any that break get tightened to assert the *correct* behavior (e.g. inline code now produces a Consolas run, not a stripped span).

### New unit tests — `tests/mcpOffice.Tests/Word/MarkdownToDocxConverterTests.cs`

Asserted against the resulting `Document` in memory (no file I/O), one per AST→docx mapping rule:

| Test | Asserts |
|---|---|
| `Headings_1_through_6_get_correct_paragraph_style` | Six headings → six paragraphs with `Heading 1`..`Heading 6` style |
| `Bold_italic_bold_italic_inline_run_flags` | `**bold**`, `*italic*`, `***both***` → three runs with the right `Bold`/`Italic` flags |
| `Inline_code_run_uses_Consolas` | `` `x` `` → run with `FontName == "Consolas"` |
| `Fenced_code_block_each_line_is_monospace_paragraph` | 3-line fence → 3 consecutive paragraphs, all Consolas, all shaded |
| `Pipe_table_creates_real_table` | 3-col × 2-row pipe table → `doc.Tables.Count == 1`, dimensions match |
| `Pipe_table_header_row_is_bold_and_shaded` | Header row cells shaded `#F2F2F2`, runs bold |
| `Pipe_table_column_alignment_from_gfm_spec` | `:---:` center column → cells centered |
| `Ordered_list_produces_numbered_paragraphs` | `1. a\n2. b` → two paragraphs in a numbered list |
| `Unordered_list_produces_bulleted_paragraphs` | `- a\n- b` → two bulleted paragraphs |
| `Nested_list_indents_per_depth` | Nested unordered → second level has greater indent |
| `Hyperlink_emits_field_with_target` | `[text](https://x)` → run inside a HYPERLINK field, target = url |
| `Image_local_file_is_embedded` | `![](rel.png)` with fixture file → `doc.Images.Count == 1` |
| `Image_missing_local_file_is_dropped_no_throw` | `![](missing.png)` → no image, no exception |
| `Image_remote_url_is_dropped_no_throw` | `![](https://x/y.png)` → no image, no exception |
| `Hr_emits_paragraph_with_bottom_border` | `---` → paragraph has bottom border |
| `Empty_markdown_produces_empty_document` | `""` → `doc.Paragraphs.Count == 1`, no content |

Image fixture is a 1×1 PNG generated programmatically at test time, not committed.

### Real-world fidelity test

Copy `C:\Projects\LimsBasic\docs\fn_send_email_callers.md` into `tests/fixtures/` (~6KB, no licensing concern), convert, assert invariants the old converter failed:

- `doc.Tables.Count >= 4` — the four category tables exist
- Body contains the literal text `"FN_SEND_EMAIL"` — proves inline code preserved
- At least one run has `Bold == true` — proves `**optional**` preserved
- Outline matches the markdown's heading tree (already covered structurally)

### Integration test

No change. The existing `Create_then_outline_via_stdio` test already covers the protocol round-trip.

## Risks

1. **DevExpress character shading on inline code may not be straightforward.** The `CharacterProperties` API exposes `BackColor` only on some surface types. Fallback: drop the gray background on inline code, keep just the Consolas font. Spike during the first inline-code test; if `BackColor` doesn't work cleanly, document the fallback in a code comment.
2. **GFM column alignment in DevExpress tables.** Per-cell horizontal alignment is set via `Cell.Range.ParagraphProperties.Alignment` — well-trodden API. Should be fine; flagging it so the alignment test verifies it.
3. **Existing tests may rely on idiosyncratic output of the old converter.** E.g. an assertion on exact paragraph count where the old converter inserted blank trailing paragraphs. Mitigation: run the old test suite first after the swap, then make minimal adjustments to assertions that were testing implementation artifacts rather than user-facing behavior.

## Out of scope

- **Syntax highlighting** in fenced code blocks. Language hint is parsed but ignored.
- **Footnotes, task lists, strikethrough, math.** Markdig parses them as literals if they appear; they survive as plain text rather than being rendered.
- **Read-back path** (`word_read_markdown`). Lossy in the other direction is a separate fix.
- **HTML embedded in markdown** (`<div>`, `<span style=…>`). Treated as literal text.
- **Network image fetching.**
- **Custom Markdig extensions** beyond `UsePipeTables()`.

## Estimated size

- New `MarkdownToDocxConverter.cs`: ~400 lines (block dispatcher + inline walker + helpers).
- Removed from `WordDocumentService.cs`: ~150 lines of regex-patching helpers.
- New tests: ~250 lines in one new file, plus the real-world fidelity test (~30 lines) in `CreateFromMarkdownTests.cs` or a new sibling file.
- Net: + ~500 lines of code, + 16 new tests, − 150 lines of regex-patching.

## Open questions deferred to implementation

1. Exact DevExpress API for character-level back colour (Risk 1). Spike during the first inline-code test.
2. Whether nested ordered+unordered list combinations (mixed) need a special path or fall out of the depth-indent logic for free.
3. Whether Markdig's `MarkdownDocument` round-trips a single trailing newline cleanly into our document (sanity-check during the empty-markdown test).
