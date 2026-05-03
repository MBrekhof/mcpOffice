# TODO

Pending work for mcpOffice. Maintained by the `/handoff` skill.

## Word POC — DONE

All 26 tasks from `docs/plans/2026-04-30-mcpoffice-word-poc-plan.md` are complete and merged (`4df3225` docs: mark Word POC final verification complete). 15 Word tools shipped.

## Excel POC

Plan: `docs/plans/2026-05-01-mcpoffice-excel-poc-design.md`. Branch: `poc/excel-tools`.

- [x] Step 1 — DevExpress Spreadsheet package references
- [x] Step 2 — Excel DTOs and `IExcelWorkbookService`
- [x] Step 3 — `excel_list_sheets`
- [x] Step 4 — `excel_read_sheet` with `maxCells`
- [x] Step 5 — Integration test for tool listing + generated-workbook read
- [x] Step 6 — Spike `excel_extract_vba` against the Air sample
- [x] Step 7 — Decision: in-process VBA extraction via `OpenMcdf` (landed)
- [x] Step 7b — `excel_extract_vba` end-to-end (synthetic unit tests + real-Excel smoke + stdio integration)
- [x] Step 8 — Formula/structure tools: `excel_get_metadata`, `excel_list_defined_names`, `excel_list_formulas`, `excel_get_structure`
- [x] Live stdio verification of `excel_extract_vba` against `C:\Projects\mcpOffice-samples\Air.xlsm` — 107 modules extracted, 566 KB payload, hasVbaProject=true. Sample dir lives outside the repo to keep ~66 MB of unrelated business workbooks out of git history. (LLM-in-the-loop verification via Claude Code MCP registration is optional follow-up — would require wiring `mcpOffice` into the MCP config and restarting Claude Code.)
- [ ] Open PR `poc/excel-tools` → `main` (squash). Suggested title: `feat: Excel POC (list_sheets, read_sheet, extract_vba, get_metadata, list_defined_names, list_formulas, get_structure)`

## Next Excel feature (post-PR)

- [ ] `excel_analyze_vba` — layer over `excel_extract_vba`. Procedures/functions with signatures, event handlers, call graph, Excel object-model references, file/DB/network deps, conversion hints. Benchmark on the 107-module `C:\Projects\mcpOffice-samples\Air.xlsm`. Strategy choice: regex-on-extracted-source vs. proper VBA tokenizer.

## Side items

### Carried from Word POC
- [ ] Wire DevExpress runtime license via `licenses.licx` once a non-trial feature is exercised.
- [ ] Optional: baseline `.editorconfig` once enough files exist to enforce against.
- [ ] Evaluate `mathieumack/MarkdownToDocxGenerator` to replace the hand-rolled markdown writer in `WordDocumentService.WriteMarkdownToDocument`. Trigger: any future task needing tables, code blocks, lists, links, or escaping. Check license + round-trip through `word_read_structured`.
- [ ] Add `[JsonDerivedType]` discriminators to the abstract `Block` record (and concrete `HeadingBlock`/`ParagraphBlock`) when tests start asserting on `word_read_structured`'s wire JSON.

### Carried from Excel POC
- [ ] Locked-VBA fixture: `tests/fixtures/sample-with-macros-locked.xlsm` to un-skip `VbaProjectReaderTests.Throws_vba_project_locked_for_protected_project`. Need a real password-protected sample to learn how Excel serializes the dir stream when locked.
- [ ] PROJECTLCID-aware code page selection in `VbaProjectReader` (currently hardcoded to cp1252). MS-OVBA dir record `0x0002 PROJECTLCID` carries the project locale.
- [ ] `excel_get_structure`: optional pivot / chart / external-connection counts via Open XML walk (DevExpress doesn't expose them directly).
- [ ] `excel_list_formulas`: rough dependency-token extraction (deferred — formula text is enough for now).
- [ ] Spike file `tests/mcpOffice.Tests/Spikes/VbaExtractionSpike.cs` is historical reference; consider removing once `excel_analyze_vba` lands.
