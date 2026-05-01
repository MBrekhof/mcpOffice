# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

mcpOffice — MCP server (stdio) exposing Office document tools, written in C# / .NET 9. Currently in POC phase: Word tools first, Excel/PowerPoint later.

Sources of truth (loaded on demand via @import):

- @SESSION_HANDOFF.md — current branch state, completed tasks, next step.
- @docs/plans/2026-04-30-mcpoffice-word-poc-plan.md — TDD task list (26 tasks, exact code for each).
- @docs/plans/2026-04-30-mcpoffice-word-poc-design.md — tool surface, error codes, design decisions.

## Build / test

- `dotnet build` — should be 0 warnings, 0 errors.
- `dotnet test` — unit + integration. Integration tests rebuild the server and spawn it via stdio (see `tests/mcpOffice.Tests.Integration/ServerHarness.cs`).
- `dotnet run --project src/mcpOffice` — runs the MCP server on stdio.

## DevExpress feed and license

- `nuget.config` references **nuget.org** plus a **local filesystem source** at `C:\Program Files\DevExpress 25.2\Components\System\Components\packages` (key `DevExpressLocal`). Local path = no URL token, no VS credential prompt. Public packages still come from nuget.org; the local source is a fallback for licensed-only packages if added later.
- Don't add `https://nuget.devexpress.com/<token>/...` URL feeds with a `%DXNUGET_KEY%` placeholder — VS prompts for credentials when the env var isn't persisted at User scope. If a remote licensed feed is truly needed, embed the token directly in the URL.
- `DevExpress_License.txt` (gitignored, repo root) — **runtime license**, the long base64 blob. Separate from any feed token. Tests currently call `RichEditDocumentServer` without an explicit license and pass (trial mode). Bake in via `licenses.licx` once non-trial features are exercised.

## MCP SDK 1.2.0 quirks

- No `IMcpClient` interface — use the concrete `McpClient` class.
- No `McpClientFactory` — use `McpClient.CreateAsync(transport)`.
- Tool names auto-lowercase unless explicit. Always set `[McpServerTool(Name = "tool_name")]`.

## Stdio discipline

stdout carries JSON-RPC. Anything written to stdout that isn't a valid JSON-RPC frame breaks the client. Logs go to **stderr only** via Serilog (already configured in `Program.cs`). Don't `Console.WriteLine` from tool code.

## Error codes

`McpException` is the only error type tools should throw. SDK 1.2.0 doesn't expose a structured `.Code` property, so codes are encoded as a `[code_string]` prefix in the message (e.g., `[file_not_found] /path/to/file.docx`). The full code list is in the design doc. Tests pattern-match on the prefix.

## Code conventions

- File-scoped namespaces, nullable enabled, implicit usings (per csproj defaults).
- Tool classes: `[McpServerToolType]` on the class, static methods with `[McpServerTool(Name=...)]` and `[Description(...)]`. See `src/mcpOffice/Tools/PingTools.cs` for the canonical shape.
- TDD: write the failing test first. Tasks 6+ in the plan have exact code for both test and implementation.
- **Test fixtures are generated programmatically** via `tests/mcpOffice.Tests/Word/TestWordDocuments.cs` (deviates from the plan's binary-fixture approach — cleaner, no committed `.docx` blobs). New Word tests should reuse this helper rather than committing `.docx` files under `tests/fixtures/`.

## Git / PRs

- Feature branch off `main` (e.g., `poc/word-tools`, `feat/<topic>`), PR back to `main`, squash merge.
- Conventional Commits: `feat:`, `fix:`, `chore:`, `test:`, `docs:`.
- Don't push to `main` directly.
