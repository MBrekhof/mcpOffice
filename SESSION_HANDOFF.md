# Session Handoff — 2026-09-02 (v4 acceptance → VBA-016, DOCS-002, VBA-011, MD-001)

## Where Things Stand

**Branch:** `main`, in sync with `origin/main`. Four PRs squash-merged today, in order: #16 `3b1f3d6` (v4 acceptance fixes + VBA-016), #17 `6093dcb` (DOCS-002), #18 `61bd8c1` (VBA-011 + PowerPoint off the roadmap), #19 `0ebe872` (MD-001), then this handoff. No side branches left, local or remote.
**Build:** `dotnet build` — 0 warnings, 0 errors. Target framework **net10.0** (SDK 10.0.400).
**Tests:** `dotnet test` — **515 unit + 21 integration pass, 2 skipped** (both smoke generators in `tests/mcpOffice.Tests/Word/MarkdownRealWorldTests.cs`). One gated Air test has a 600 ms performance budget and flakes when a build runs alongside it — rerun before believing it.
**Tool surface:** **38 tools**: 1 ping + 15 Word + 15 Excel + 7 PDF. No new tools today; `excel_map_vba_sheet_access` gained `includeRecords` and `maxRecords`.

## What Landed Today (2026-09-02)

1. **Live acceptance of the v4 tools** on the samples corpus through the office server — the item the previous handoff left open. Three of four passed on the first run; every finding below became a card and a fix.
   - `excel_list_vba_entry_points(Air.xlsm)` — 233 entry points, 39 unreachable, one genuinely unresolved form-control macro (`StartDiscreteAnalyzer` on sheet `no3+no2`, no such procedure). Bug: a formula calling a UDF three times listed the cell three times.
   - `excel_list_vba_form_controls(OlieGC - LABWARE PRD.xlsm)` — 3 forms, 6 controls. Bug: `Label2_Click` typed CommandButton because VBE default names had no prefix entry.
   - `excel_compare_vba_corpus(directory)` — 20 workbooks, 79 shared procedures, one call, no timeout. Groupings plausible (the three `kalibratieberekening` books share `frmSerialInput`/`Module6`; the two Mediaformulier books are copies).
   - `excel_map_vba_sheet_access(Air.xlsm)` unscoped — 114 KB on one line, over the client's tool-result cap; the caller saw only "output saved to file".
2. **VBA-012 / VBA-015 fixes** (`c489aa1`, PR #16) — distinct `formulaCells`; the MSForms type names (`Label2`, `TextBox1`, …) count as prefixes, confidence `prefix`.
3. **VBA-016** (`50a05dc` + `d45823d`, PR #16) — `excel_map_vba_sheet_access(…, includeRecords=true, maxRecords=100)`. `includeRecords=false` returns the summary and the per-sheet rollup only (9 KB on Air) — the first call on a big workbook. Measured live: 300 records = 59 KB still overflowed Claude Code's cap, hence 100 (~26 KB with the rollup, then seen live). Design doc "Tool 2 — Output" records the numbers.
4. **DOCS-002** (`af37ce6`, PR #17) — `word_convert` documented as taking `.md`/`.markdown` input (tool description, README, usage.md); behaviour was already there.
5. **VBA-011** (`40f1f46`, PR #18) — the stranded LF-only renderer pass from `feat/render-vba-callgraph` applied verbatim via `git diff | git apply`; one no-`\r` test per renderer. The local branch is deleted.
6. **PowerPoint dropped from the roadmap** (PR #18) — README, CLAUDE.md, handoff. User decision; no consumer asked for it.
7. **MD-001** (`c9c82fb`, PR #19) — one `WriteInline` for paragraphs and table cells behind an `InsertionPoint` anchor (−38 lines). Behaviour change: images are inserted at the insertion point (`Images.Insert`, verified in dxdocs 26.1) instead of appended at the document end, so an image inside a table cell is kept where it used to be dropped. Test `Image_inside_a_table_cell_is_inserted_in_that_cell`.

## Outstanding — Action Required

- **Board:** Confirm Done in the UI for VBA-006, VBA-007, VBA-011, VBA-012, VBA-013, VBA-014, VBA-015, VBA-016, DOCS-002, MD-001 (plus MD-003, DOCS-001, WORD-001 from the previous session if still in Review). Every conclusion cites its commit and PR.
- **Other machine:** `git pull`, restart its session so the server picks up the new DLL.
- **`/mcp`** in this session if the office server is wanted — it was killed for the last build.

## Next Up

Board is the source of truth (`list_cards`, project id 27). Nothing is claimed; the previous shortlist is empty. Candidates by size:

- **VBA-009** — scanner tests for `ParamArray` and `Static Sub` forms (small).
- **VBA-010** — pagination on `callGraph` / `references` in `excel_analyze_vba`: the payload problem VBA-016 just solved for sheet access, same `maxRecords` + `truncated` shape.
- **CHORE-001** — baseline `.editorconfig`.
- **PDF-001** — `pdf_extract_tables` on top of `LineGrouper` (the largest open feature).

## How To Resume

```powershell
cd C:\Projects\mcpOffice
git pull
git log --oneline -8
dotnet build --nologo
dotnet test --nologo
```

## Operational notes

- The MCP server holds a lock on `bin\Debug\net10.0\mcpOffice.dll` (MSB3027). Kill and build in the
  *same* command: `pid=$(powershell -NoProfile -Command "(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'dotnet.exe' -and $_.CommandLine -like '*mcpOffice.dll*' }).ProcessId"); [ -n "$pid" ] && taskkill //PID $pid //F //T; dotnet build --nologo`.
  Then `/mcp` in Claude Code to respawn the server against the fresh DLL. Every build in a session
  costs one `/mcp`; batch builds accordingly.
- **Live acceptance is the bar, and it finds things unit tests cannot**: the 114 KB payload was
  invisible to a gated test that only checked "≤ cap". When a tool returns a list, check the live
  call lands in the client before calling the card done.
- **Implementer subagents: one at a time.** Two agents running `dotnet test` in the same tree collide
  on obj/bin and the DLL lock. The lead should not build while an agent is building either.
- A subagent was killed mid-task by the account's session limit ("resets 3:50am"); it had written
  nothing. Check `git status` for partial files before assuming an agent's report is complete.
