# TODO

Pending work for mcpOffice. Maintained by the `/handoff` skill.

## Word POC — DONE

All 26 tasks from `docs/plans/2026-04-30-mcpoffice-word-poc-plan.md` are complete and merged (`4df3225` docs: mark Word POC final verification complete). 15 Word tools shipped.

## Excel POC — DONE

Plan: `docs/plans/2026-05-01-mcpoffice-excel-poc-design.md`. All 8 steps shipped across PRs #1, #2, #3 (squash-merged into `main`; feature branches deleted). 7 Excel tools on main: `excel_list_sheets`, `excel_read_sheet`, `excel_extract_vba`, `excel_get_metadata`, `excel_list_defined_names`, `excel_list_formulas`, `excel_get_structure`. Live stdio verification against the real 107-module `C:\Projects\mcpOffice-samples\Air.xlsm` confirmed end-to-end correctness. (LLM-in-the-loop verification via Claude Code MCP registration is optional and unblocked.)

## Next Excel feature (post-PR)

- [x] `excel_analyze_vba` — DONE (branch `feat/excel-analyze-vba`, merged). Procedures/functions with signatures, event handlers, call graph, Excel object-model references, file/DB/network deps. Benchmarked against the 107-module `C:\Projects\mcpOffice-samples\Air.xlsm`: 200 procedures, 110 event handlers, 938 call edges, 3040 object-model reference sites, 48 external dependencies, ~115ms wall time.

## excel_analyze_vba v2 — conversion hints layer

These items are surfaced by the v1 Air.xlsm benchmark as the natural next step toward Excel-to-C# migration tooling:

- [ ] Conversion hints per procedure: classify as event handler / utility / data-transform / UI glue; suggest C# equivalent (method, service class, hosted service, etc.).
- [ ] Dependency graph rendering: emit a DOT/Mermaid call graph for agent consumption.
- [ ] Cross-module coupling score: identify tightly coupled module clusters as refactoring targets.
- [ ] `excel_analyze_vba` v2 design doc — capture the shape of conversion hints DTO before implementing.

## Side items

### Carried from Word POC
- [ ] Optional: baseline `.editorconfig` once enough files exist to enforce against.
- [ ] Add `[JsonDerivedType]` discriminators to the abstract `Block` record (and concrete `HeadingBlock`/`ParagraphBlock`) when tests start asserting on `word_read_structured`'s wire JSON.

### Carried from Excel POC
- [ ] Locked-VBA fixture: `tests/fixtures/sample-with-macros-locked.xlsm` to un-skip `VbaProjectReaderTests.Throws_vba_project_locked_for_protected_project`. Need a real password-protected sample to learn how Excel serializes the dir stream when locked.
- [ ] PROJECTLCID-aware code page selection in `VbaProjectReader` (currently hardcoded to cp1252). MS-OVBA dir record `0x0002 PROJECTLCID` carries the project locale.
- [ ] `excel_get_structure`: optional pivot / chart / external-connection counts via Open XML walk (DevExpress doesn't expose them directly).
- [ ] `excel_list_formulas`: rough dependency-token extraction (deferred — formula text is enough for now).
- [ ] Synthetic extract→analyze integration test using `VbaProjectBinBuilder` — design doc promised this; deferred in v1 in favor of the Air.xlsm gated benchmark. Adds coverage on machines without the sample file.
- [x] `VbaProjectReader.ClassifyKind` locale-dependent heuristic. DONE. `Read(xlsmPath)` now extracts the OOXML codenames from `xl/workbook.xml` (`workbookPr/codeName`) and every sheet xml (`worksheets/`, `chartsheets/`, `dialogsheets/` → `sheetPr/codeName`) into a set, which is passed into `ReadVbaProjectBin` and on into `ClassifyKind`. When the set is non-null, classification is purely by membership — locale-independent and survives user-renamed codenames. The legacy English-prefix heuristic remains as fallback for callers that don't pass codenames (e.g. synthetic `VbaProjectBinBuilder` tests). Verified against `RingOnderzoek.xlsm` (Dutch — `Blad1`/`Blad3` now `documentModule`) and `Balans.xlsm` (Dutch — `Blad3` now `documentModule`); Air.xlsm regression-clean.
- [ ] `VbaProcedureScanner` lacks tests for `ParamArray` parameter form and `Static Sub` procedure form. Both currently parse correctly per the regex; no behavior gap, just test coverage.
- [x] `excel_get_structure` parse_error on `RingOnderzoek.xlsm`. DONE. Root cause: DevExpress.Spreadsheet `WorksheetCollection` on this workbook is internally inconsistent — `Count` returns 1, `foreach` yields 0, and `Worksheets[0]` throws. Fix: introduced `MaterializeWorksheets()` helper that enumerates via foreach (which works fine on healthy files) and applied it in `ListSheets`, `GetStructure`, and `ResolveWorksheet`. Service now returns whatever can be enumerated rather than throwing — degenerate workbooks land as `sheetCount: 0, sheets: []`. Two watchdog tests in `RingOnderzoekStructureTests` will fail when DevExpress fixes their indexer, signalling the workaround can be removed.
- [x] **`excel_analyze_vba` per-module filter.** DONE. `moduleName` parameter added to the tool, service interface, and analyzer. Case-insensitive match; throws `module_not_found` (with available names listed) when unknown; null/empty preserves whole-workbook output. Summary stays whole-workbook so the caller still has accurate totals; the `modules`, `callGraph`, and `references` arrays are filtered to entries involving the focal module (call edges include both directions: from-module and resolved-into-module). `sheetName` was dropped from scope — sheets aren't the natural axis in VBA-land; codename↔sheet-name mapping is a separate, harder feature that nobody has asked for.
- [ ] **Pagination on `callGraph` and `references` arrays in `excel_analyze_vba`.** Even with a `moduleName` filter, the heaviest module on a large workbook can be too big. Add `offset` / `limit` (or cursor) to the heavy arrays so a caller can stream them. Lower priority than the module filter — the filter alone covers most real cases.
