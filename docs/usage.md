# mcpOffice Usage

## Requirements

- .NET 10 SDK
- DevExpress 26.1 installed locally
- DevExpress license file kept outside source control

This repo currently restores DevExpress packages from the local offline package source installed at:

```text
C:\Program Files\DevExpress 26.1\Components\System\Components\packages
```

## Build And Test

```powershell
dotnet restore
dotnet build --no-restore
dotnet test --no-restore
```

Expected current test count: 134 passing tests / 1 skipped (the locked-VBA fixture, awaiting a real password-protected sample).

## Run The MCP Server

For local development:

```powershell
dotnet run --project C:\Projects\mcpOffice\src\mcpOffice
```

VS Code workspace config is already committed at `.vscode/mcp.json`. It starts the Debug build with:

```json
{
  "servers": {
    "office": {
      "type": "stdio",
      "command": "dotnet",
      "args": [
        "${workspaceFolder}/src/mcpOffice/bin/Debug/net10.0/mcpOffice.dll"
      ]
    }
  }
}
```

Claude Code config is committed at `.mcp.json` at the repo root. It uses the same Debug DLL but with an **absolute** path, since Claude Code does not expand `${workspaceFolder}`-style variables:

```json
{
  "mcpServers": {
    "office": {
      "command": "dotnet",
      "args": [
        "C:\\Projects\\mcpOffice\\src\\mcpOffice\\bin\\Debug\\net10.0\\mcpOffice.dll"
      ]
    }
  }
}
```

If your checkout lives at a different path, edit the `args` value before launching Claude Code in this directory. Restart Claude Code after creating or editing `.mcp.json` so it reloads the MCP server list.

Run `dotnet build` before starting either server so the DLL exists.

For release/client configuration, publish first:

```powershell
dotnet publish C:\Projects\mcpOffice\src\mcpOffice -c Release -r win-x64 --self-contained false
```

The published executable is created under:

```text
C:\Projects\mcpOffice\src\mcpOffice\bin\Release\net10.0\win-x64\publish\mcpOffice.exe
```

Generic MCP client entry for the published executable:

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\Projects\\mcpOffice\\src\\mcpOffice\\bin\\Release\\net10.0\\win-x64\\publish\\mcpOffice.exe"
    }
  }
}
```

## Available Tools

Read tools:

- `Ping`: returns `pong`.
- `word_get_outline(path)`: returns heading nodes from a `.docx`.
- `word_get_metadata(path)`: returns core properties, page count, and word count.
- `word_read_markdown(path)`: returns a conservative Markdown projection.
- `word_read_structured(path)`: returns headings, paragraphs with runs, tables, images, and properties.
- `word_list_comments(path)`: returns comment summaries.
- `word_list_revisions(path)`: returns tracked-change summaries.

Write/create tools:

- `word_create_blank(path, overwrite=false)`: creates an empty `.docx`.
- `word_create_from_markdown(path, markdown, overwrite=false, templatePath?)`: creates `.docx` from Markdown. With `templatePath` (absolute path to a `.dotx` or `.docx`) the document is built on that template: its Normal and Heading 1-6 styles, headers/footers and page setup carry over, and the Markdown body is appended after any content the template already contains. Where the template defines a heading style, the template's formatting wins over the converter's built-in heading look; heading levels the template does not define fall back to the converter defaults.
- `word_append_markdown(path, markdown)`: appends Markdown to an existing `.docx`.
- `word_find_replace(path, find, replace, useRegex=false, matchCase=false)`: replaces text and returns replacement count.
- `word_insert_paragraph(path, atIndex, text, style?)`: inserts a paragraph.
- `word_insert_table(path, atIndex, headers[], rows[][])`: inserts a table.
- `word_set_metadata(path, properties)`: sets `author`, `title`, `subject`, and/or `keywords`.
- `word_mail_merge(templatePath, outputPath, dataJson, overwrite=false)`: replaces `{{token}}` placeholders; pass `overwrite=true` to regenerate into an existing output path.

Convert tools:

- `word_convert(inputPath, outputPath, format?, overwrite=false)`: converts a `.docx` or a `.md`/`.markdown` file to `pdf`, `html`, `rtf`, `txt`, `md`/`markdown`, or `docx`. If `format` is omitted, it is inferred from `outputPath`. Pass `overwrite=true` to regenerate into an existing output path. A Markdown input is rendered by the same engine as `word_create_from_markdown`, so for a Markdown file that already exists on disk this is the route — `report.md` → `report.pdf` in one call, nothing passes through the agent's context; relative image paths resolve against the `.md`'s directory.

Excel read tools:

- `excel_list_sheets(path)`: returns sheets with visibility, used range, dimensions.
- `excel_read_sheet(path, sheetName?, sheetIndex?, range?, includeFormulas=true, includeFormats=false, maxCells=50000)`: returns rows + addressed cell details for a worksheet or A1 range.
- `excel_get_metadata(path)`: returns workbook document properties + sheet count.
- `excel_list_defined_names(path)`: returns workbook + sheet-scoped names with `refersTo`, `comment`, `isHidden`.
- `excel_list_formulas(path, sheetName?, includeValues=false, maxFormulas=10000)`: returns formula cells with optional cached values; raises `range_too_large` when capped.
- `excel_get_structure(path, includeSheets=true, includeFormulaCounts=true, includeDefinedNames=true)`: returns a workbook-level rollup with optional per-sheet detail.

Excel macro tools:

- `excel_extract_vba(path)`: returns raw VBA module source from `.xlsm` (in-process via OpenMcdf — no Excel install required). For `.xlsx` or workbooks without macros, returns `hasVbaProject=false`.
- `excel_analyze_vba(path, includeProcedures=true, includeCallGraph=false, includeReferences=false)`: layered structural analysis on top of `excel_extract_vba` — procedures with signatures, event handlers, FQN-resolved call graph, Excel object-model references with literal-arg capture, and filesystem/database/network/automation/shell dependency dispatch. Tiered output via toggles.
- `excel_render_vba_callgraph(path, format="mermaid", moduleName?, procedureName?, depth=2, direction="both", layout="clustered", maxNodes=300)`: the call graph as Mermaid or DOT; narrow with `moduleName` / `procedureName` on large workbooks or hit `graph_too_large`.
- `excel_suggest_vba_conversion(path, moduleName?, targetParadigm?)`: per-procedure conversion hints (trigger / purity / shape / dependencies, rationale, and a C# emission target when `targetParadigm` is `classLibrary | workerService | webApi | console`) plus whole-workbook module coupling.
- `excel_list_vba_entry_points(path, includeUnreachable=true, moduleName?)`: what actually runs — `eventHandler`, `autoMacro`, `shapeMacro` and `formControlMacro` (macros wired to shapes and buttons, read from the drawing parts), `worksheetFunction` (Public Functions used in cell formulas), `dynamicDispatch` (`Application.OnTime` / `OnKey` / `Run`, `.OnAction`, `CallByName` with literal targets) — and `unreachable[]`, the procedures no entry point can reach (confidence `high` | `medium`). The migration scope cut.
- `excel_map_vba_sheet_access(path, moduleName?, sheetName?, includeUnresolved=true, includeRecords=true, maxRecords=100)`: per procedure, which sheet and range / defined name it reads and writes (`mode` read | write | both), with a per-sheet readers/writers rollup. On a big workbook call with `includeRecords=false` first (summary + rollup only, a few KB), then scope the records with `moduleName` / `sheetName`; `maxRecords` caps them with `truncated: true`. Resolves sheet names, `Sheets(n)`, codenames (`Blad1.Range`), a sheet module's own unqualified `Range`/`Cells`, `With` blocks, one-assignment aliases and defined names; `ActiveSheet` and unqualified access elsewhere are returned as unresolved with a reason, never guessed.
- `excel_compare_vba_corpus(paths[]? | directory, minOccurrences=2, maxProcedures=200, includeNearDuplicates=true)`: procedures shared across workbooks — `identical` (same normalised body; renamed copies group) and `nearDuplicate` (same name, ≥ 90 % similar) — plus `sharedModules[]`. Loads every workbook's VBA project; per-file failures land in `workbooks[].error`.
- `excel_list_vba_form_controls(path, formName?)`: each UserForm's controls inferred from its code-behind — handler names (`cmdOK_Click`), `Me.<control>` references, Hungarian-prefixed or VBE-default-named bare references (`txtName`, `Label2`) and `As MSForms.<Type>` declarations — with `inferredType`, `typeConfidence` (`declared | prefix | event | member | none`), events and referenced properties, plus the form's own `formEvents`. The binary `.frx` designer part is not read.

PDF read tools:

- `pdf_get_metadata(path)`: returns document properties, PDF version, permission flags, bookmark count, and per-page `{pageNumber, width, height, rotation}` in points (1/72 inch).
- `pdf_read_text(path, pageRange?, preserveLayout=false, maxChars=200000)`: returns text per page. With `preserveLayout=true` the page is rebuilt as a fixed-width grid — the equivalent of `pdftotext -layout` — which is what keeps a column report readable; leave it off for prose.
- `pdf_read_layout(path, pageRange?, granularity="line", includeFontInfo=false, maxWords=50000)`: returns positioned text. `granularity="word"` gives one entry per word, `"line"` groups words into visual lines.
- `pdf_find_text(path, query, caseSensitive=false, wholeWords=false, maxResults=500)`: returns every match with page and bounding box.
- `pdf_get_outline(path)`: returns the bookmark tree; empty when the PDF has none.

PDF export tools:

- `pdf_render_page(path, pageNumber, outputPath, dpi=150, format?, overwrite=false)`: renders one page to `png`, `jpg`/`jpeg`, `bmp`, `gif`, or `tiff`. Format is inferred from `outputPath` when omitted.
- `pdf_extract_images(path, outputDirectory, pageRange?, minPixelSize=16, maxImages=200, overwrite=false)`: writes embedded raster images as `page{NNN}_img{NNN}.png` and reports where each sat on the page.

**`pageRange`** accepts a comma-separated list of 1-based pages and inclusive spans: `"1"`, `"2-5"`,
`"1,3,7-9"`, `"5-"` (to the last page). Omit it for every page.

**PDF coordinates** are reported in points with the origin at the **top-left** of the page, so
sorting by `y` ascending gives reading order. PDF's own coordinate system has the origin at the
bottom-left; the flip is done for you.

There is no PDF *writer* here — `word_convert` already produces PDFs from `.docx` and `.md`.

All `path`, `inputPath`, `outputPath`, `outputDirectory`, and `templatePath` values must be absolute Windows paths.

## Example Calls

Create from Markdown:

```json
{
  "path": "C:\\Temp\\proposal.docx",
  "markdown": "# Proposal\n\nHello **world**.\n\n| Name | Value |\n| ---- | ----- |\n| Alpha | 1 |",
  "overwrite": false
}
```

Read outline:

```json
{
  "path": "C:\\Temp\\proposal.docx"
}
```

Expected `word_get_outline` shape:

```json
[
  { "level": 1, "text": "Proposal" }
]
```

Convert to PDF:

```json
{
  "inputPath": "C:\\Temp\\proposal.docx",
  "outputPath": "C:\\Temp\\proposal.pdf"
}
```

Mail merge:

```json
{
  "templatePath": "C:\\Temp\\template.docx",
  "outputPath": "C:\\Temp\\merged.docx",
  "dataJson": "{\"firstName\":\"Ada\",\"score\":42}"
}
```

Extract VBA from a macro-enabled workbook:

```json
{
  "path": "C:\\Workbooks\\AnalysisTool.xlsm"
}
```

Analyze VBA structure (cheap procedure list only):

```json
{
  "path": "C:\\Workbooks\\AnalysisTool.xlsm"
}
```

Analyze VBA with full call graph and references (heaviest output):

```json
{
  "path": "C:\\Workbooks\\AnalysisTool.xlsm",
  "includeProcedures": true,
  "includeCallGraph": true,
  "includeReferences": true
}
```

Expected `word_get_metadata` shape:

```json
{
  "author": "Martin",
  "title": "Proposal",
  "subject": "MCP Office",
  "keywords": "mcp,office,word",
  "created": "2026-04-30T10:00:00",
  "modified": "2026-04-30T11:00:00",
  "lastPrinted": null,
  "revisionCount": 7,
  "pageCount": 1,
  "wordCount": 1200
}
```

## Markdown Notes

`word_create_from_markdown` and `word_append_markdown` use a Markdig-based converter (`MarkdownToDocxConverter`). Current coverage: headings, paragraphs, bold/italic, nested lists, block quotes, horizontal rules, fenced and indented code blocks, inline code, links, local images, and pipe tables. `word_create_from_markdown` can build on a `.dotx`/`.docx` template via `templatePath` (see above).

Known caveats:

- Lists currently round-trip through `word_read_structured` as paragraph text with literal `-` or `1.` prefixes, not semantic Word list objects.
- Hyperlink URLs are not exposed by `word_read_structured` yet.
- Markdown export (`word_read_markdown` and `.md` conversion) is a conservative projection, not a full-fidelity Markdown serializer.

## Error Codes

Tool errors are returned as `McpException` messages prefixed with stable codes:

- `[file_not_found]`
- `[file_exists]`
- `[invalid_path]`
- `[unsupported_format]`
- `[parse_error]`
- `[index_out_of_range]`
- `[merge_field_missing]`
- `[io_error]`
- `[internal_error]`
- `[sheet_not_found]` — Excel: named sheet not in workbook
- `[range_too_large]` — Excel: result would exceed `maxCells` / `maxFormulas`
- `[vba_project_missing]` — reserved for future strict mode of `excel_extract_vba`
- `[vba_project_locked]` — VBA project is password-protected for viewing
- `[vba_parse_error]` — OLE walk / MS-OVBA decompression / dir-record-walk failure
- `[password_required]` — PDF: the document is encrypted
- `[page_not_found]` — PDF: the requested page is outside the document
- `[invalid_page_range]` — PDF: `pageRange` could not be parsed
- `[invalid_render_option]` — PDF: `dpi` outside 12-1200 (also used by the callgraph renderer)

## Troubleshooting

- If restore fails with `NU1301 The local source '...' doesn't exist`, the DevExpress feed path in `nuget.config` no longer matches the installed version. Point it at `C:\Program Files\DevExpress <major>\Components\System\Components\packages` and bump the `DevExpress.*` package versions in `src/mcpOffice` **and** `tests/mcpOffice.Tests` to match. Nothing in the repo compiles until both are aligned.
- If VS Code cannot start the MCP server, run `dotnet build` and confirm `src\mcpOffice\bin\Debug\net10.0\mcpOffice.dll` exists.
- If tool calls fail with `[invalid_path]`, pass an absolute path such as `C:\Docs\file.docx`.
- If tool calls fail with `[file_not_found]`, confirm the MCP server process can access the file.
- If output calls fail with `[file_exists]`, use a different output path or pass `overwrite=true` where the tool supports it.
- If the MCP client hangs, verify the server logs go to stderr only; stdout is reserved for the MCP JSON-RPC stream.
