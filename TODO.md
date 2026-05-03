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
- [ ] **Actionable now:** Remove spike file `tests/mcpOffice.Tests/Spikes/VbaExtractionSpike.cs` — it was kept as historical reference pending `excel_analyze_vba` landing; that has now landed. Production code (`VbaProjectReader`, `VbaSourceAnalyzer`, et al.) is fully independent of the spike.
