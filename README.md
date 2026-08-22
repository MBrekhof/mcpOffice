# mcpOffice

An MCP (Model Context Protocol) server for Microsoft Office documents, written in C# (.NET 10) and backed by DevExpress Office File API packages. It lets AI agents read, write, and convert Office documents through tool calls instead of one-off scripts.

**Status:** Word (.docx), Excel (.xlsx / .xlsm) and PDF are shipped — 34 tools. Excel includes `excel_analyze_vba` v3 (procedures, event handlers, call graph, object-model references, external dependencies, conversion hints). PDF covers metadata, text with layout preservation, positioned words/lines, search, page rendering, embedded images and bookmarks. Next: PowerPoint (.pptx).

## Architecture

![mcpOffice architecture](docs/img/architecture.svg)

Source: [`docs/img/architecture.excalidraw`](docs/img/architecture.excalidraw) (open in [Excalidraw](https://excalidraw.com)). See [ARCHITECTURE.md](ARCHITECTURE.md) for the full layer map.

## Documents

- [Architecture](ARCHITECTURE.md) — layer map, domains, tool-adding pattern, error model, VBA pipeline and PDF text-positioning diagrams.
- [Usage](docs/usage.md) — build, run, MCP client config, sample calls, troubleshooting.
- [Word design](docs/plans/2026-04-30-mcpoffice-word-poc-design.md) — Word tool surface, error model.
- [Word implementation plan](docs/plans/2026-04-30-mcpoffice-word-poc-plan.md) — task-by-task TDD plan.
- [Markdown-to-docx design](docs/plans/2026-05-07-mcpoffice-markdown-to-docx-markdig-design.md) — Markdig-based converter behind `word_create_from_markdown` / `word_append_markdown`.
- [Excel design](docs/plans/2026-05-01-mcpoffice-excel-poc-design.md) — Excel tool surface and rationale.
- [VBA extraction plan](docs/plans/2026-05-01-mcpoffice-excel-vba-extraction-plan.md) — MS-OVBA decompression, OpenMcdf walking.
- [VBA analysis design](docs/plans/2026-05-03-mcpoffice-excel-analyze-vba-design.md) and [v3 conversion hints](docs/plans/2026-05-07-mcpoffice-excel-analyze-vba-v3-design.md).
- [CSV export design](docs/plans/2026-05-07-mcpoffice-excel-export-csv-design.md) — streaming `excel_export_csv`.
- [PDF tools design](docs/plans/2026-08-20-pdf-tools-design.md) — `pdf_` surface, top-left coordinate convention, layout reconstruction.

## Current Tools

34 tools shipped: 1 ping + 15 Word + 11 Excel + 7 PDF.

### Word

- `word_get_outline(path)`
- `word_get_metadata(path)`
- `word_read_markdown(path)`
- `word_read_structured(path)`
- `word_list_comments(path)`
- `word_list_revisions(path)`
- `word_create_blank(path, overwrite=false)`
- `word_create_from_markdown(path, markdown, overwrite=false)`
- `word_append_markdown(path, markdown)`
- `word_find_replace(path, find, replace, useRegex=false, matchCase=false)`
- `word_insert_paragraph(path, atIndex, text, style?)`
- `word_insert_table(path, atIndex, headers[], rows[][])`
- `word_set_metadata(path, properties)`
- `word_mail_merge(templatePath, outputPath, dataJson)`
- `word_convert(inputPath, outputPath, format?)`

### Excel

- `excel_list_sheets(path)` — sheets in order with visibility, used range, dimensions.
- `excel_read_sheet(path, sheetName?, sheetIndex?, range?, includeFormulas=true, includeFormats=false, maxCells=50000)` — cell data with formulas + formats.
- `excel_get_metadata(path)` — author, title, created/modified, sheet count, document properties.
- `excel_list_defined_names(path)` — workbook + sheet-scoped names with refersTo / scope / hidden flag.
- `excel_list_formulas(path, sheetName?, includeValues=false, maxFormulas=10000)` — formula cells with optional cached values.
- `excel_get_structure(path, includeSheets=true, includeFormulas=true, includeDefinedNames=true)` — workbook rollup sized for huge workbooks.
- `excel_extract_vba(path)` — static VBA module source via in-process MS-OVBA decompression (no Excel install required).
- `excel_analyze_vba(path, includeProcedures=true, includeCallGraph=false, includeReferences=false, moduleName?)` — structural analysis on top of the extracted source: procedures with signatures, event handlers, call graph (with intra-workbook resolution), Excel object-model references, and external dependencies (file/DB/network/automation/shell). Pass `moduleName` to scope the heavy arrays to a single module on large workbooks; the summary stays whole-workbook.
- `excel_export_csv(path, outputPath, sheetName?, sheetIndex?, range?, overwrite=false, maxRows=1048576, trimTrailingEmptyRows=false)` — streams a sheet to RFC 4180 CSV for pandas/polars.
- `excel_render_vba_callgraph(path, format="mermaid", moduleName?, procedureName?, depth=2, direction="both", layout="clustered", maxNodes=300)` — call graph as Mermaid or DOT.
- `excel_suggest_vba_conversion(path, moduleName?, targetParadigm?)` — per-procedure conversion hints plus module coupling.

### PDF

- `pdf_get_metadata(path)` — title/author/subject/keywords/creator/producer, dates, PDF version, permission flags, bookmark count, and per-page width/height/rotation in points.
- `pdf_read_text(path, pageRange?, preserveLayout=false, maxChars=200000)` — text per page. `preserveLayout=true` rebuilds the fixed-width grid (like `pdftotext -layout`) so column reports stay aligned and sliceable by character position.
- `pdf_read_layout(path, pageRange?, granularity="line", includeFontInfo=false, maxWords=50000)` — positioned text: every word or visual line with x/y/width/height, **origin top-left** so sorting by y ascending is reading order.
- `pdf_find_text(path, query, caseSensitive=false, wholeWords=false, maxResults=500)` — every match with its page and bounding box.
- `pdf_render_page(path, pageNumber, outputPath, dpi=150, format?, overwrite=false)` — render a page to png/jpg/bmp/gif/tiff, for scanned PDFs or when extracted text is ambiguous.
- `pdf_extract_images(path, outputDirectory, pageRange?, minPixelSize=16, maxImages=200, overwrite=false)` — embedded raster images to PNG, with their placement on the page.
- `pdf_get_outline(path)` — bookmark tree as nested `{title, level, pageNumber, children}`.

`pageRange` accepts `"1"`, `"2-5"`, `"1,3,7-9"` and `"5-"` (to the end); omit it for every page.

### Other

- `Ping` — health check, returns `pong`.

All file paths passed to tools must be absolute.

## Example

Create a Word document from Markdown, then convert it to PDF:

```json
{
  "path": "C:\\Temp\\proposal.docx",
  "markdown": "# Proposal\n\nHello **Word**.",
  "overwrite": false
}
```

```json
{
  "inputPath": "C:\\Temp\\proposal.docx",
  "outputPath": "C:\\Temp\\proposal.pdf"
}
```

Extract VBA modules from a macro-enabled workbook:

```json
{
  "path": "C:\\Workbooks\\AnalysisTool.xlsm"
}
```

Read a column-based report out of a PDF with its layout intact:

```json
{
  "path": "C:\\Reports\\overzichtsrapport.pdf",
  "pageRange": "1-3",
  "preserveLayout": true
}
```

## Roadmap

1. **Word POC** — read / write / convert .docx ✓
2. **Excel POC** — read sheets, list formulas/structure/defined names, extract VBA ✓
3. **`excel_analyze_vba` v1** — call graph, event handlers, Excel object-model refs, external dependencies ✓
4. **`excel_analyze_vba` v2** — conversion hints (procedure role classification, suggested C# equivalents, DOT/Mermaid call-graph rendering, cross-module coupling score).
5. **PDF** — metadata, text (with layout preservation), positioned words/lines, search, page rendering, embedded images, bookmarks ✓
6. PowerPoint (.pptx).

## Built With

- [`ModelContextProtocol`](https://github.com/modelcontextprotocol/csharp-sdk) — C# MCP SDK.
- DevExpress `Document.Processor` 26.1 — server-side RichEdit / Spreadsheet / PDF APIs.
- [`Markdig`](https://github.com/xoofx/markdig) — Markdown parser behind the Markdown-to-DOCX converter.
- [`OpenMcdf`](https://www.nuget.org/packages/OpenMcdf) — OLE compound file reader for VBA project extraction.
