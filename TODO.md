# TODO

Pending work for mcpOffice. Maintained by the `/handoff` skill.

## Word POC — DONE

All 26 tasks from `docs/plans/2026-04-30-mcpoffice-word-poc-plan.md` are complete and merged (`4df3225` docs: mark Word POC final verification complete). 15 Word tools shipped.

## Excel POC — DONE

Plan: `docs/plans/2026-05-01-mcpoffice-excel-poc-design.md`. All 8 steps shipped across PRs #1, #2, #3 (squash-merged into `main`; feature branches deleted). 7 Excel tools on main: `excel_list_sheets`, `excel_read_sheet`, `excel_extract_vba`, `excel_get_metadata`, `excel_list_defined_names`, `excel_list_formulas`, `excel_get_structure`. Live stdio verification against the real 107-module `C:\Projects\mcpOffice-samples\Air.xlsm` confirmed end-to-end correctness. (LLM-in-the-loop verification via Claude Code MCP registration is optional and unblocked.)

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
