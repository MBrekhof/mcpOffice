# Session Handoff — 2026-11-08 — Excel write tools merged, deployed and live-accepted

## Where Things Stand

**Branch:** `main`, which includes #28 (Excel write tools) and #29 (culture fix). There are no open PRs or feature branches.
**Build:** `dotnet build` gives 0 warnings, 0 errors. Target framework net10.0, DevExpress **26.1.6**. `dotnet list src\mcpOffice package --vulnerable --include-transitive` is clean.
**Tests:** **706 unit + 24 integration pass, 2 skipped** (the Markdown fixture generators).
**Tool surface:** **40 tools**, including `excel_create_workbook` and `excel_update_workbook`.

## What Landed

| Commit | PR | What |
|--------|----|------|
| `12d05ad` | #28 | `excel_create_workbook` / `excel_update_workbook` (#2853). Design: `docs/plans/2026-11-04-mcpoffice-excel-write-design.md`. Two Gemini reviews found 14 issues; each is fixed and has its own test in `Excel/Write/WriteEdgeCaseTests.cs`. |
| (squash) | #29 | `fontSize` and column-`width` validation messages put a double straight into the message text, which follows the machine's culture, so nl-NL produced `fontSize 0,5`. `WriteEdgeCaseTests.FormatRange_fontSize_and_borderScope_validation` failed every time on a Dutch machine. The fix is `FormattableString.Invariant`, plus a regression test pinned to nl-NL. |

**Recorded design deviation:** a formula *written by an update* is evaluated at write time. Without that, DevExpress caches `0` for it. Existing cached values are left alone, and `FullCalculationOnLoad` is still set.

## Share Deployment — done 2026-11-05

`\\dev-san-001\applicatiedata\WLN\Labware\Claude\tools\mcpOffice` now holds a self-contained win-x64 Release built from `775fd66` (the #29 branch head). Its code is identical to the squash-merged `main`. That means DevExpress 26.1.6, 40 tools, and #2854 shipped.

- **Before the copy:** the old share (`f7aa8b2`, 26.1.4) was backed up to the session scratchpad. That backup is temporary.
- **Smoke test** of the published exe over stdio: `initialize`, then `tools/list` (40 tools), then `excel_create_workbook`, then `excel_read_sheet` round-tripped. Dates came back typed and `=SUM` returned 19.5.
- **robocopy** reported no `*EXTRA` files, because the runtime is unchanged at 10.0.14.
- **Hashes** of `mcpOffice.dll`, `mcpOffice.exe` and `DevExpress.Docs.v26.1.dll` match the publish output.

How to deploy:
1. `dotnet publish src\mcpOffice -c Release -r win-x64 --self-contained true -o <scratch>`.
2. Smoke-test the exe.
3. Back up the share.
4. `robocopy <scratch> <share> /E`.
5. Remove any `*EXTRA` version-named runtime files.
6. Verify the hashes.

robocopy exit codes 1–3 mean success.

## Board

| Card | Column | State |
|------|--------|-------|
| #2853 "Create Excel files" | Review | Merged in #28, deployed, live-accepted 2026-11-08 (conclusion records it). Martin confirms Done by hand. |
| #2854 vulnerable crypto packages | Todo → **Review (move by hand; agents can't)** | The conclusion records the fix (#26), the clean vulnerability scan and the redeploy. |

## Outstanding

- **Live acceptance of the write tools: done 2026-11-08** through the live `office` MCP server. A
  2-sheet `.xlsx` (dates, booleans, evaluated `=SUM`, formatted total row, merged cell) and a new
  version of corpus `bevestigingslijst.xlsm` (added sheet + formula). The VBA project and all three
  form-control button macros on sheet `monsters` survived the DevExpress round trip; only the shape
  names changed (`_x0000_s1029` → `Button_x0020_5`). `warnings[]` empty on both. Both opened in
  Excel and accepted.
- **`~/.claude.json` user-scope `office` entry** pointed at the dead `net9.0` build and conflicted
  with the project `.mcp.json` ("Conflicting scopes"). Repointed to `net10.0` on 2026-11-08; a
  `net9.0` folder still exists under `src/mcpOffice/bin/Debug/` and can be deleted.
- **Known flaky test:** `AirSampleConversionHintsTests.Suggests_conversion_for_air_workbook_within_budget` asserts under 600 ms. It once took 685 ms on a loaded machine.

## Next Up

These candidates are unchanged:
- **WORD-004** (1468): the `word_read_markdown` payload cap. First answer the `_meta: anthropic/maxResultSizeChars` question for SDK 1.2.0.
- **WORD-005** (1469): body list markers in the Markdown projection.
- **VBA-009**: scanner tests.
- **VBA-010**: pagination.
- **CHORE-001**: `.editorconfig`.
- **PDF-001**: `pdf_extract_tables`.

## The Tool-Result Cap (still applies to WORD-004)

- Claude Code caps MCP results **client-side at 25,000 tokens**, with a warning at 10,000. It *reports characters but enforces tokens*.
- The full result always crosses stdio, so unit tests can't see an overflow.
- **Do not propose `MAX_MCP_OUTPUT_TOKENS`.** It was rejected on policy on 2026-09-02.
- Scoping fixes context cost; raising the cap doesn't.

## How To Resume

```powershell
cd C:\Projects\mcpOffice
git pull
dotnet build --nologo
dotnet test --nologo
```

## Operational Notes

- **The machine culture is nl-NL.** Any number put into a message or test string has to be formatted invariant, or tests pass only on English machines. #29 was the first bite.
- **The server holds a lock on `bin\Debug\net10.0\mcpOffice.dll` (MSB3027).** Kill it and build in the *same* command:
  `$p = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'dotnet.exe' -and $_.CommandLine -like '*mcpOffice.dll*' }; if ($p) { $p | ForEach-Object { taskkill /PID $_.ProcessId /F /T | Out-Null } }; dotnet build --nologo`.
  Then run `/mcp`. The session that builds can't bring the server back itself.
- **Smoke-testing over stdio by hand:** the server handles requests **concurrently**, so replies can come back out of order. Send a request and read its reply before sending the next. Keep stdin open until the replies have arrived; on EOF the server exits before it flushes.
- **PowerShell 5.1 splits `git commit -m` here-strings that contain double quotes.** The commit silently fails and the push still runs. Use `git commit -F <file>` instead.
- **Gemini review:** `agy -p "$(cat prompt)" --model gemini-4-argon-high --mode plan --sandbox --dangerously-skip-permissions --output-format stream-json --print-timeout 40m > out.jsonl`. The review text is in the last line's `result.response`.
- **`gh` account:** `MartinWLN` is the default and can't create PRs on `MBrekhof/mcpOffice`. Run `gh auth switch --hostname github.com --user MBrekhof`, then switch back.
- **DevExpress installer 26.1.6.** If restore fails with `NU1301`, the feed folder moved: repoint `nuget.config` and the `DevExpress.*` versions together (CLAUDE.md). New NuGet audit warnings can appear without any code change; try the newest DevExpress patch before pinning packages.
- **Editing a csproj with `[IO.File]::WriteAllText` drops its UTF-8 BOM.** Use the Edit tool.
- **The ODT benchmark** is `C:\Projects\WLNCentral\rewab\20221220 Handleiding Risicogestuurd monitoren.odt`.
