# TODO

Pending work for mcpOffice. Maintained by the `/handoff` skill.

## Word POC — DONE

All 26 tasks from `docs/plans/2026-04-30-mcpoffice-word-poc-plan.md` are complete and merged (`4df3225` docs: mark Word POC final verification complete). 15 Word tools shipped.

## Excel POC — DONE

Plan: `docs/plans/2026-05-01-mcpoffice-excel-poc-design.md`. All 8 steps shipped across PRs #1, #2, #3 (squash-merged into `main`; feature branches deleted). 7 Excel tools on main: `excel_list_sheets`, `excel_read_sheet`, `excel_extract_vba`, `excel_get_metadata`, `excel_list_defined_names`, `excel_list_formulas`, `excel_get_structure`. Live stdio verification against the real 107-module `C:\Projects\mcpOffice-samples\Air.xlsm` confirmed end-to-end correctness.

## Excel analyzer v1 + v2 — DONE

- [x] **`excel_analyze_vba` (v1)** — DONE (PR #4, merged). Procedures/functions with signatures, event handlers, call graph, Excel object-model references, file/DB/network deps. Benchmarked against the 107-module `Air.xlsm`: 200 procedures, 110 event handlers, 938 call edges, 3040 object-model reference sites, 48 external dependencies, ~115ms wall time.
- [x] **`excel_render_vba_callgraph` (v2)** — DONE (PR #12, `feat/render-vba-callgraph`, squash-merged as `f93831c`). New 25th MCP tool that renders the VBA call graph as Mermaid (default) or DOT. Layered on `excel_analyze_vba`; the analyzer is unchanged. New `VbaCallgraphFilter` (pure function): whole-workbook / `moduleName` direct-neighbour / focal-procedure BFS with `depth` and `direction`. `MermaidCallgraphRenderer` + `DotCallgraphRenderer` behind `ICallgraphRenderer`. New error codes: `procedure_not_found`, `graph_too_large`, `invalid_render_option`. Verified against Air.xlsm: whole-workbook render trips `graph_too_large`; single-module render succeeds; focal-BFS depth=1 < 500ms. Supersedes the stale PR #9 (closed without merge — predated #10/#11).

## excel_analyze_vba v3 — conversion-hints layer (still pending)

These remain as the natural next step toward Excel-to-C# migration tooling once the visualization layer is in:

- [ ] Conversion hints per procedure: classify as event handler / utility / data-transform / UI glue; suggest C# equivalent (method, service class, hosted service, etc.).
- [ ] Cross-module coupling score: identify tightly coupled module clusters as refactoring targets.
- [ ] v3 design doc — capture the shape of conversion-hints DTO before implementing. Use `docs/plans/2026-05-03-mcpoffice-excel-render-vba-callgraph-design.md` as the shape template.

## Word md→docx fidelity — Markdig converter (branch ready for PR)

- [x] **Replace lossy `MarkdownToDocxGenerator` with Markdig AST walker.** DONE (branch `feat/markdown-to-docx-markdig`, 22 commits, NOT yet merged). `MarkdownToDocxConverter` handles paragraphs, headings 1–6, ordered/unordered/nested lists, fenced + indented code blocks, blockquotes, thematic breaks, GFM tables (bold+shaded header, column alignment), bold/italic/bold-italic, inline code (Consolas), hyperlinks, autolinks, hard+soft line breaks, local image embed, remote image drop. Affects `word_create_from_markdown`, `word_append_markdown`, `word_convert` (.md input). Real-world fidelity verified against `fn_send_email_callers.md` (4+ tables, inline code, bold). 206 unit + 13 integration green.
- [x] **Table cell inline formatting.** DONE (commit on `feat/markdown-to-docx-markdig`). `CollectCellText` removed; `WriteTable` now uses a `CellCursor` + `WriteCellInline` that anchors each inline write to the live `dxCell.ContentRange` so backtick code (Consolas), bold, italic, hyperlinks and line breaks inside table cells all render with their proper formatting. Root cause was that `doc.Paragraphs.Get(cellContentRange)` returns stale paragraph positions in table cells — fixed by re-reading `dxCell.ContentRange.Start` fresh for each cell and tracking the cursor forward through insertions. New test: `Table_cells_render_inline_formatting`.

## Side items

### Carried from Word POC
- [ ] Optional: baseline `.editorconfig` once enough files exist to enforce against.
- [ ] Add `[JsonDerivedType]` discriminators to the abstract `Block` record (and concrete `HeadingBlock`/`ParagraphBlock`) when tests start asserting on `word_read_structured`'s wire JSON.

### Carried from Excel POC
- [ ] PROJECTLCID-aware code page selection in `VbaProjectReader` (currently hardcoded to cp1252). MS-OVBA dir record `0x0002 PROJECTLCID` carries the project locale.
- [ ] `excel_get_structure`: optional pivot / chart / external-connection counts via Open XML walk (DevExpress doesn't expose them directly).
- [ ] `excel_list_formulas`: rough dependency-token extraction (deferred — formula text is enough for now).
- [x] **DevExpress formula parser/serializer leaks host culture (nl-NL).** DONE (PR #11, `fix/devexpress-defined-name-refersto`, squash-merged as `6175f4d`). Both `ExcelWorkbookService.LoadWorkbook` (read side) and `TestExcelWorkbooks.Create` (test fixture write side) now set `Workbook.Options.Culture = CultureInfo.InvariantCulture`. Fixes the nl-NL failure where `DefinedNames.Add("TaxRate", "=0.21")` threw `ArgumentException` (DevExpress parsed `0.21` as `0` + invalid `.21` because `,` is the decimal separator in nl-NL) and `ListDefinedNames` returned `RefersTo` as `=0,21` to the agent. MCP API now serves locale-neutral formula text regardless of host locale.
- [x] Synthetic extract→analyze integration test. DONE (PR #10, `feat/synthetic-analyze-test`, squash-merged as `128f5bd`). `tests/mcpOffice.Tests/Excel/Vba/SyntheticAnalyzeTests.cs` runs unconditionally against `tests/fixtures/synthetic-vba.xlsm` (Excel-authored via `tests/fixtures/Generate-SyntheticVbaXlsm.ps1`). Replaces the originally-planned `VbaProjectBinBuilder` route with a real Excel-authored fixture so the test exercises Excel's actual MS-OVBA copy-token compressed chunks (the synthetic builder only emits literal-only chunks). Asserts 4-module structure, ParamArray + Static-Sub forms parse, locale-agnostic `documentModule` classification (Dutch `Blad1` codename), event-handler classification, cross-module call edge `ThisWorkbook.Workbook_Open → Module1.Main`, and Excel object-model refs.
- [x] `VbaProjectReader.ClassifyKind` locale-dependent heuristic. DONE. `Read(xlsmPath)` now extracts the OOXML codenames from `xl/workbook.xml` (`workbookPr/codeName`) and every sheet xml (`worksheets/`, `chartsheets/`, `dialogsheets/` → `sheetPr/codeName`) into a set, which is passed into `ReadVbaProjectBin` and on into `ClassifyKind`. When the set is non-null, classification is purely by membership — locale-independent and survives user-renamed codenames. The legacy English-prefix heuristic remains as fallback for callers that don't pass codenames (e.g. synthetic `VbaProjectBinBuilder` tests). Verified against `RingOnderzoek.xlsm` (Dutch — `Blad1`/`Blad3` now `documentModule`) and `Balans.xlsm` (Dutch — `Blad3` now `documentModule`); Air.xlsm regression-clean.
- [ ] `VbaProcedureScanner` lacks tests for `ParamArray` parameter form and `Static Sub` procedure form. Both currently parse correctly per the regex; no behavior gap, just test coverage.
- [x] `excel_get_structure` parse_error on `RingOnderzoek.xlsm`. DONE. Root cause: DevExpress.Spreadsheet `WorksheetCollection` on this workbook is internally inconsistent — `Count` returns 1, `foreach` yields 0, and `Worksheets[0]` throws. Fix: introduced `MaterializeWorksheets()` helper that enumerates via foreach (which works fine on healthy files) and applied it in `ListSheets`, `GetStructure`, and `ResolveWorksheet`. Service now returns whatever can be enumerated rather than throwing — degenerate workbooks land as `sheetCount: 0, sheets: []`. Two watchdog tests in `RingOnderzoekStructureTests` will fail when DevExpress fixes their indexer, signalling the workaround can be removed.
- [x] **`excel_analyze_vba` per-module filter.** DONE. `moduleName` parameter added to the tool, service interface, and analyzer. Case-insensitive match; throws `module_not_found` (with available names listed) when unknown; null/empty preserves whole-workbook output. Summary stays whole-workbook so the caller still has accurate totals; the `modules`, `callGraph`, and `references` arrays are filtered to entries involving the focal module (call edges include both directions: from-module and resolved-into-module). `sheetName` was dropped from scope — sheets aren't the natural axis in VBA-land; codename↔sheet-name mapping is a separate, harder feature that nobody has asked for.
- [ ] **Pagination on `callGraph` and `references` arrays in `excel_analyze_vba`.** Even with a `moduleName` filter, the heaviest module on a large workbook can be too big. Add `offset` / `limit` (or cursor) to the heavy arrays so a caller can stream them. Lower priority than the module filter — the filter alone covers most real cases.
- [ ] **`excel_export_csv` (pandas-style bulk export).** `excel_read_sheet` is fine for slicing but caps at 50k cells and returns a JSON cell-grid — wrong shape for "load this sheet as a dataframe" workflows. Stream a sheet (or A1 range) to CSV on disk and return the output path, so the agent can hand it to `pandas.read_csv` / `polars.read_csv` instead of reassembling JSON pages. Open questions: header-row option, date format (default ISO 8601), decimal/thousand separators (invariant culture), formula cells (value vs formula text), `na_rep` for blanks, `.csv.gz` for big sheets. Likely also wants an `excel_export_ndjson` sibling for column-typed output.
