# TODO

Pending work for mcpOffice. Pulled from `docs/plans/2026-04-30-mcpoffice-word-poc-plan.md` plus session-level items. Maintained by the `/handoff` skill.

## Plan tasks

- [x] Task 1 — Initialize git repo and baseline files (.gitignore, README, nuget.config)
- [x] Task 2 — Create solution and three projects
- [x] Task 3 — Add NuGet package references
- [x] Task 4 — Minimal Program.cs with stdio MCP host + ping tool
- [x] Task 5 — Integration test that spawns server and calls ping
- [x] Task 6 — Error model: `McpToolException` with stable codes (`ErrorCode.cs`, `ToolError.cs`)
- [x] Task 7 — `PathGuard` (absolute-path / file-existence / writable checks)
- [x] Task 8 — Tool implementation pattern + first Word tool (`word_get_outline`)
- [x] Task 9 — `word_get_metadata`
- [x] Task 10 — `word_read_markdown`
- [ ] Task 11 — `word_read_structured`  ← next
- [ ] Task 12 — `word_list_comments`
- [ ] Task 13 — `word_list_revisions`
- [ ] Task 14 — `word_create_blank`
- [ ] Task 15 — `word_create_from_markdown`
- [ ] Task 16 — `word_append_markdown`
- [ ] Task 17 — `word_find_replace`
- [ ] Task 18 — `word_insert_paragraph`
- [ ] Task 19 — `word_insert_table`
- [ ] Task 20 — `word_set_metadata`
- [ ] Task 21 — `word_mail_merge`
- [ ] Task 22 — `word_convert`
- [ ] Task 23 — Tool-surface integration test (partial — `ToolSurfaceTests.cs` exists but lists only the 3 implemented tools; expand as new tools land)
- [ ] Task 24 — End-to-end integration test per tool group
- [ ] Task 25 — Docs: `docs/usage.md` (exists, may need polish) + README polish
- [ ] Task 26 — Final verification (Release build, publish, live MCP wire-in)

## Side items

- [ ] Wire DevExpress runtime license via `licenses.licx` once a non-trial feature is exercised. Currently all `RichEditDocumentServer` calls succeed under trial mode.
- [ ] Optional: add a baseline `.editorconfig` once a few more files exist, so `dotnet format` has rules to enforce.
- [ ] Decide whether `tests/mcpOffice.Tests/Word/TestWordDocuments.cs` (programmatic fixture builder) should replace the plan's binary `tests/fixtures/*.docx` approach for all remaining Word tasks. Current code commits to programmatic — plan tasks 11-21 should follow suit.
