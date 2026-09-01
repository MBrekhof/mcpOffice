# Session Handoff — 2026-08-22 (sync, docs refresh, .NET 10, OpenMcdf fix)

## Where Things Stand

**Branch:** `main` — clean working tree, in sync with `origin/main`.
**Build:** `dotnet build` — 0 warnings, 0 errors. Target framework **net10.0** (SDK 10.0.400).
**Tests:** `dotnet test` — **351 unit + 17 integration pass, 2 skipped** (both smoke generators in `tests/mcpOffice.Tests/Word/MarkdownRealWorldTests.cs`; the locked-VBA skip was deleted in d026685).
**Tool surface:** **34 tools**: 1 ping + 15 Word + 11 Excel + 7 PDF.

The PDF tools branch from 2026-08-20 is merged — it landed on `main` as `6d83594`. There is no open feature branch.

## What Landed This Session

1. **Synced this machine** — pulled six commits from the other machine (Word converter fixes, `templatePath`, DevExpress 26.1 realignment, the seven `pdf_` tools).
2. **Docs refresh** (`4609ac8`) — architecture diagram gained the PDF column (`docs/img/architecture.{excalidraw,png,svg}`, re-exported via the Excalidraw canvas); README status / documents / Built With brought current; stray `DevExpress 25.2` references in `ARCHITECTURE.md` and `docs/usage.md` fixed.
3. **OpenMcdf 3.1.3 → 3.1.4** in both `src/mcpOffice` and `tests/mcpOffice.Tests`. The NU1902 the previous handoff called "pre-existing" was GHSA-5qwm-7pvp-w988: an *uncatchable infinite loop* on a crafted CFB directory cycle — a malicious `.xlsm` could hang the server with nothing for the `try/catch` wrapper to catch. Patched in 3.1.4.
4. **net9.0 → net10.0** across all three projects, `ServerHarness.cs`, `.mcp.json`, `docs/usage.md`. `System.Text.Encoding.CodePages` dropped from both csprojs — it is part of the shared framework on .NET 10 (NU1510). cp1252 VBA decoding still passes.

## Outstanding — Action Required

- **`.mcp.json` now points at `bin\Debug\net10.0\mcpOffice.dll`.** Claude Code reads `.mcp.json` at session start, so the `office` server needs a **session restart** (not just `/mcp`) on each machine to pick up the new path. Until then clients see whatever DLL the old path still holds.
- **Other machine:** `git pull` before working. Its clone still has `TODO.md`; until it pulls, the sync hook prints the BOARD-ONLY refusal — that's the guard working, not an error to fix by re-adding the file.

## Next Up

**Task state is board-only as of this session.** ContextBoard project `mcpOffice` (id 27): `TODO.md`
deleted, `FileSyncClosed` set on the VPS (refusal proven: `fileSyncClosed: true, updated: 0`), 25
backlog cards minted from the file's items — VBA-001…010, CSV-001…004, MD-001/002, PDF-001…006,
WORD-001/002, EXCEL-001/002, CHORE-001. `list_cards` is the source of truth; the `/handoff` skill
was rewritten to reconcile the board instead of the file. Shortlist by card:

- **WORD-001** — `overwrite` parameter on `word_mail_merge` (Todo, 0.5h; the only card not in Backlog).
- **CSV-001** `excel_export_ndjson`, **CSV-002** `.csv.gz`.
- **VBA-006** `file → filesystem` schema drift and **VBA-007** `mdl` module prefix — the two real bugs in the backlog.
- **MD-001** `WriteCellInline` / `WriteInline` unification, **MD-002** Normal-style polish.
- **PowerPoint (.pptx)** — next domain per the README roadmap; no design doc and no card yet.

## How To Resume

```powershell
cd C:\Projects\mcpOffice
git pull
git log --oneline -5
dotnet build --nologo
dotnet test --nologo
```

## Operational note

The MCP server picks up new code only when its process restarts, and while running it holds a lock
on `bin\Debug\net10.0\mcpOffice.dll` that fails the build with `MSB3027`. Pattern that works:

1. `Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like "*mcpOffice.dll*" }` — find PID.
2. `taskkill //PID <pid> //F //T` — release the lock.
3. `dotnet build src/mcpOffice --nologo` — rebuild Debug (the registered MCP path).
4. `/mcp` in Claude Code — reconnect, which respawns the server against the fresh DLL.

Claude Code respawns the server automatically after a disconnect, so it can retake the lock between
your kill and your build — kill and build in the *same* command. This session's live server was a
Debug DLL from **2026-05-13** (27 tools, no PDF) until step 2 — check the DLL timestamp when tools
seem to be missing.
