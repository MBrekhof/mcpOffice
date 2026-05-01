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
- [x] Task 11 — `word_read_structured`
- [x] Task 12 — `word_list_comments`
- [x] Task 13 — `word_list_revisions`
- [x] Task 14 — `word_create_blank`
- [x] Task 15 — `word_create_from_markdown`
- [x] Task 16 — `word_append_markdown`
- [x] Task 17 — `word_find_replace`
- [x] Task 18 — `word_insert_paragraph`
- [x] Task 19 — `word_insert_table`
- [x] Task 20 — `word_set_metadata`
- [x] Task 21 — `word_mail_merge`
- [ ] Task 22 — `word_convert`  ← next
- [ ] Task 23 — Tool-surface integration test (`ToolSurfaceTests.cs` already lists all 15 current tools; remaining work is to lock the spec once Task 22 lands)
- [ ] Task 24 — End-to-end integration test per tool group
- [ ] Task 25 — Docs: `docs/usage.md` (exists, may need polish) + README polish
- [ ] Task 26 — Final verification (Release build, publish, live MCP wire-in)

## Side items

- [ ] Wire DevExpress runtime license via `licenses.licx` once a non-trial feature is exercised. Currently all `RichEditDocumentServer` calls succeed under trial mode.
- [ ] Optional: add a baseline `.editorconfig` once a few more files exist, so `dotnet format` has rules to enforce.
- [ ] Decide whether `tests/mcpOffice.Tests/Word/TestWordDocuments.cs` (programmatic fixture builder) should replace the plan's binary `tests/fixtures/*.docx` approach for all remaining Word tasks. Current code commits to programmatic — plan tasks 11-21 followed suit; 22+ should too.
- [ ] Evaluate replacing the hand-rolled markdown writer in `WordDocumentService.WriteMarkdownToDocument` with [`mathieumack/MarkdownToDocxGenerator`](https://github.com/mathieumack/MarkdownToDocxGenerator). Trigger: any future task that needs tables, code blocks, lists, links, or escaping in the markdown create/append path. Check license compatibility and whether output round-trips through `word_read_structured`.
- [ ] Add `[JsonDerivedType]` discriminators to the abstract `Block` record (and concrete `HeadingBlock`/`ParagraphBlock`) once Task 24 starts asserting on `word_read_structured`'s wire JSON. Unit tests don't need it.
