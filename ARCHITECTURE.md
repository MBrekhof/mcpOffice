# Architecture

A landing-page map of mcpOffice. Stable content only — counts and status live in `SESSION_HANDOFF.md`; open work lives on ContextBoard (project `mcpOffice`). Per-feature design lives in `docs/plans/`.

## What it is

A stdio MCP server (.NET 10, C#) that exposes Office-document operations as tools. Stateless: every tool call takes an absolute file path, performs one operation, returns. No session handles, no open-document lifecycle for the agent to manage.

## Layers

```
JSON-RPC over stdio
        |
        v
   Program.cs                 -- host bootstrap; AddMcpServer + WithToolsFromAssembly
        |
        v
   Tools/*Tools.cs            -- thin [McpServerTool] facade; one class per domain
        |
        v
   Services/<Domain>/         -- IXxxService + impl; the real work
        |                       (PathGuard, parse, transform, save)
        v
   DevExpress.Document.Processor
   (RichEditDocumentServer, SpreadsheetControl)
```

Cross-cutting:

- `Models/` — DTOs returned to the agent (records).
- `ErrorCode.cs` + `ToolError.cs` — stable string error codes encoded as `[code] message`.
- `PathGuard.cs` — absolute-path / exists / writable preconditions.

## Domains

| Domain  | Tool class            | Service                     | Tool prefix |
|---------|-----------------------|-----------------------------|-------------|
| Health  | `Tools/PingTools`     | —                           | (none)      |
| Word    | `Tools/WordTools`     | `Services/Word/WordDocumentService` | `word_`     |
| Excel   | `Tools/ExcelTools`    | `Services/Excel/ExcelWorkbookService` | `excel_`    |
| PDF     | `Tools/PdfTools`      | `Services/Pdf/PdfDocumentService`   | `pdf_`      |

The Excel VBA pipeline lives under `Services/Excel/Vba/` and has more moving parts than the rest of the codebase — see "VBA pipeline" below. The PDF domain has its own wrinkle — see "PDF text positioning" below.

## Adding a new tool

The pattern, end-to-end (TDD):

1. **Decide domain.** Reuse existing `*Tools` + `*Service` if the tool fits a domain that already exists.
2. **Add DTOs** under `Models/` if the return shape is new. Records, file-scoped namespace, nullable enabled.
3. **Extend the service interface** (`IXxxService`) and impl. Service methods do path guarding, DevExpress work, and translate any non-`McpException` to `ToolError.ParseError(...)` (or a more specific code) inside a `try/catch (Exception ex) when (ex is not McpException)` wrapper.
4. **Write the failing unit test** under `tests/mcpOffice.Tests/<Domain>/`. Generate fixtures programmatically — see `Word/TestWordDocuments.cs` and `Excel/TestExcelWorkbooks.cs`. Don't commit `.docx`/`.xlsm` blobs unless there's a reason a programmatic fixture won't do.
5. **Implement** until the test passes.
6. **Add the `[McpServerTool(Name = "<domain>_<verb>")]`** method on the `*Tools` class. Always set `Name` explicitly (SDK lowercases otherwise). Keep tool methods as one-liners that delegate to the service.
7. **Update `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`** with the new tool name. This is the canary that catches accidental tool removal/rename.
8. **Conventional Commit**, feature branch off `main`, PR, squash-merge.

**Design for an LLM caller.** The consumer is an agent with a context window, not a UI. Scope before paginate (a `moduleName`-style filter beats a page cursor). Tiered detail — heavy payloads opt-in. Every record self-contained (`fullyQualifiedName`, never an index into a sibling array). Closed enum vocabularies written down in the design doc. Renderers return a bare string.

## Error model

Tools throw `McpException` only. The SDK doesn't expose a structured `Code` property in 1.2.0, so codes are encoded as a `[code]` prefix in the message:

```
[file_not_found] C:\path\to\missing.docx
```

Tests pattern-match on the prefix. Codes are stable identifiers an agent can branch on; the rest of the message is human-readable and may change. Codes live in `ErrorCode.cs`; helpers that throw them live in `ToolError.cs`.

## Stdio discipline

stdout is the JSON-RPC channel — anything written to it that isn't a valid frame breaks the client. Logging goes to **stderr only** via Serilog (configured in `Program.cs`). Never `Console.WriteLine` from tool or service code.

## Tests

Two suites:

- **`tests/mcpOffice.Tests/`** — unit tests against services directly. No transport. Fast. The bulk of correctness coverage lives here.
- **`tests/mcpOffice.Tests.Integration/`** — spawns the built server exe via stdio and round-trips through `McpClient`. Kept tiny on purpose (each test boots a process). Covers the tool catalog (`ToolSurfaceTests`), one happy path per workflow (`WordWorkflowTests`, `ExcelWorkflowTests`), and the ping smoke test.

Fixtures: programmatic generators in `TestWordDocuments.cs` / `TestExcelWorkbooks.cs`. Hand-authored binaries land in `tests/fixtures/` only when programmatic generation isn't viable (e.g. macro-enabled `.xlsm` with a real VBA project — DevExpress can't author one).

A gated "real-world" benchmark exists for VBA analysis: `Excel/Vba/AirSampleAnalysisTests.cs` runs against `C:\Projects\mcpOffice-samples\Air.xlsm` and skips when the file is absent.

## VBA pipeline (Excel)

`excel_extract_vba` and `excel_analyze_vba` share a layered pipeline under `Services/Excel/Vba/`:

```
.xlsm  -- (Open XML, DevExpress) -->  vbaProject.bin
                                          |
                                          v
                              MsOvbaDecompressor  (MS-OVBA RLE)
                                          |
                                          v
                              VbaDirStreamParser  (module records)
                                          |
                                          v
                              VbaProjectReader    (module sources, cp1252)
                                          |
                                          +--->  excel_extract_vba returns here
                                          |
                                          v
                              VbaSourceAnalyzer   (per-module orchestrator)
                                |        |        |
                                v        v        v
                              Cleaner  Scanner  CallGraph + ReferenceCollector
                                          |
                                          v
                              excel_analyze_vba returns the structural model
```

Strategy is regex-on-cleaned-source rather than a full VBA tokenizer. The Air.xlsm benchmark (107 modules, 938 call edges, 3040 object-model sites, ~115ms) is the evidence that this is sufficient. Revisit only if real-world ambiguity defeats the regex layer.

The v4 migration-planning tools (2026-09) compose that model with the package itself instead of
DevExpress — `OpenXmlParts` maps sheets ↔ codenames ↔ drawing parts and streams formulas and
defined names straight out of the zip, because DevExpress Spreadsheet exposes form controls but
not the macro they run:

```
excel_list_vba_entry_points   DrawingMacroExtractor (drawingN.xml, vmlDrawingN.vml)
                              + VbaDynamicDispatchScanner + formulas  -->  entry points
                              VbaCallGraphReachability over the v1 graph -->  unreachable[]
excel_map_vba_sheet_access    VbaSheetAccessResolver (With / alias / codename / defined-name
                              aware, never guesses ActiveSheet)         -->  per-sheet read/write map
excel_compare_vba_corpus      VbaProcedureHasher (normalised body, name excluded) across
                              N workbooks                                -->  shared procedures
```

Each of those is a pure class with string-fixture tests; the service layer only opens the zip and
hands the parts over. Design and closed vocabularies:
`docs/plans/2026-09-01-mcpoffice-excel-vba-v4-migration-planning-design.md`.

## PDF text positioning

A PDF is a list of glyph-drawing operations. It has no lines, no columns, and no reading
order — those are things a renderer infers. Three small pure classes under `Services/Pdf/`
own that inference, which is why they are unit-tested independently of any PDF:

```
PdfDocumentProcessor.NextWord()   -- forward-only cursor over the whole document
        |                            (no per-page overload; unwanted pages are walked
        v                             and dropped, still one pass)
   PdfWordBox[]                   -- text + box, Y FLIPPED to a top-left origin so
        |                            "sort by Y ascending" is reading order
        v
   LineGrouper                    -- clusters words whose vertical centres fall within
        |                            0.6 x word height; tolerance scales with font size
        v
   PdfTextLine[]                  -- pdf_read_layout granularity="line" returns these
        |
        v
   LayoutTextRenderer             -- places each word at round(x / medianCharWidth),
                                     i.e. pdftotext -layout, for pdf_read_text
                                     preserveLayout=true
```

The flip matters: DevExpress reports `Top` in PDF user space, where Y grows *upward* from the
bottom of the page. Every DTO in `Models/Pdf*` documents the top-left convention.

`preserveLayout` exists because `GetPageText` returns words in content-stream order with no
horizontal padding, which collapses column-based reports — lab result sheets, invoices,
anything originally printed as monospaced text — into ambiguous runs.

## DevExpress

- Package: `DevExpress.Document.Processor` (server-side, no UI). `RichEditDocumentServer` for Word, `SpreadsheetControl`/Open XML walks for Excel, `PdfDocumentProcessor` for PDF.
- `PdfDocumentProcessor` lives in `DevExpress.Docs.vXX.dll` (the Document.Processor package), **not** in `DevExpress.Pdf.Core` — that package holds the model types (`PdfDocument`, `PdfPage`, `PdfWord`, `PdfOrientedRectangle`).
- Runtime license: `DevExpress_License.txt` at the repo root, gitignored.
- NuGet: nuget.org + a local filesystem source at `C:\Program Files\DevExpress 26.1\...\packages` (key `DevExpressLocal` in `nuget.config`). No URL feed with a token — that prompts for credentials in VS.

## What this architecture deliberately does not do

- No CI workflow yet — manual `dotnet build` / `dotnet test`.
- No multi-platform builds — Windows-only because of DevExpress runtime.
- No installer / signing.
- No retry on file locks — surfaces `io_error`, agent retries.
- No auth / sandboxing — runs locally as the user, full filesystem access.

## Where to look for what

| You want to know...                           | Look at                                              |
|-----------------------------------------------|------------------------------------------------------|
| Current branch state, last session            | `SESSION_HANDOFF.md`                                 |
| What's pending                                | ContextBoard, project `mcpOffice` (`list_cards`)     |
| Why a feature is shaped the way it is         | `docs/plans/<date>-<feature>-design.md`              |
| How a feature was built                       | `docs/plans/<date>-<feature>-plan.md`                |
| How to wire the server into Claude Code       | `docs/usage.md`                                      |
| Project conventions, MCP SDK quirks, stdio    | `CLAUDE.md`                                          |
| The shape of the codebase (this file)         | `ARCHITECTURE.md`                                    |
