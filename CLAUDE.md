# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

mcpOffice — MCP server (stdio) exposing Office document tools, written in C# / .NET 10. Word, Excel and PDF domains are shipped; PowerPoint is next.

Sources of truth (loaded on demand via @import):

- @ARCHITECTURE.md — layer map, domains, tool-adding pattern, error model, VBA pipeline.
- @SESSION_HANDOFF.md — current branch state, completed tasks, next step.
- **Open work lives on ContextBoard** (project `mcpOffice`, id 27) — `list_cards` / `get_card`. This repo is **board-only** since 2026-08-22: never create `TODO.md` or `DOCS/DONE.md` (the server refuses file-sync pushes for this project).
- @docs/plans/2026-04-30-mcpoffice-word-poc-plan.md — TDD task list (26 tasks, exact code for each).
- @docs/plans/2026-04-30-mcpoffice-word-poc-design.md — tool surface, error codes, design decisions.

## Build / test

- `dotnet build` — should be 0 warnings, 0 errors.
- `dotnet test` — unit + integration. Integration tests rebuild the server and spawn it via stdio (see `tests/mcpOffice.Tests.Integration/ServerHarness.cs`).
- `dotnet run --project src/mcpOffice` — runs the MCP server on stdio.
- **Acceptance = the live `office` MCP server in this session on a real file** (corpus: `C:\Projects\mcpOffice-samples`), not a harness one-shot or a Python script. Rebuild + `/mcp` per the SESSION_HANDOFF operational note first. Reaching for Python or a scratch harness means a tool is missing or awkward — card that gap, don't work around it.

## DevExpress feed and license

- `nuget.config` references **nuget.org** plus a **local filesystem source** at `C:\Program Files\DevExpress 26.1\Components\System\Components\packages` (key `DevExpressLocal`). Local path = no URL token, no VS credential prompt. Public packages still come from nuget.org; the local source is a fallback for licensed-only packages if added later.
- **The feed path is version-pinned and NuGet local sources take no wildcard.** Upgrading the DevExpress installer deletes the old folder, and then *every* restore in the repo dies with `NU1301 the local source '...' doesn't exist` — nothing compiles, including tests. When that happens, repoint `nuget.config` **and** the `DevExpress.*` `PackageReference` versions (in `src/mcpOffice` *and* `tests/mcpOffice.Tests`) to the installed major in the same commit. This bit the repo once already: it sat on 25.2.5 after the machine moved to 26.1 (fixed 2026-08-20).
- Packages resolve out of the installer's **fallback folder** (`C:\Program Files\DevExpress 26.1\Components\Offline Packages`), not `~/.nuget/packages` — so don't conclude a DevExpress package is missing just because the global cache has no folder for it.
- Don't add `https://nuget.devexpress.com/<token>/...` URL feeds with a `%DXNUGET_KEY%` placeholder — VS prompts for credentials when the env var isn't persisted at User scope. If a remote licensed feed is truly needed, embed the token directly in the URL.
- `DevExpress_License.txt` (gitignored, repo root) — **runtime license**, the long base64 blob. Separate from any feed token. Tests currently call `RichEditDocumentServer` without an explicit license and pass (trial mode). Bake in via `licenses.licx` once non-trial features are exercised.
- **RichEdit has no Markdown format** (dxdocs 26.1 `DocumentFormat`: Doc/Docx/Dot*/ePub/FlatOpc*/Html/Mht/Odt/OpenXml/PlainText/Rtf/WordML). The Word POC plan doc's Tasks 10/15/22 (`DocumentFormat.Markdown`, `Options.Export.Markdown`) are wrong — md→docx is Markdig + `MarkdownToDocxConverter`, docx→md is the hand-rolled `RenderMarkdown`.

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
- Tests use xUnit `Assert.*` only. FluentAssertions was removed (8.x is commercial); the plan docs under `docs/plans/` still show FluentAssertions code — translate it, never re-add the package.
- **RichEdit inherits everything.** Every paragraph, table cell and run appended through `RichEditDocumentServer` inherits the previous paragraph style, list index, direct paragraph formatting and the previous run's character properties. Five bleed bugs so far, MD-003 is the sixth. Reset at the single append point in `MarkdownToDocxConverter` (`InsertRun`); add a bleed test for every new Markdig node type. Font names live in per-script slots — `Reset(CharacterPropertiesMask.FontName)` alone does nothing visible, reset every `FontName*` mask (that is why the July fix was a no-op).

## Git / PRs

- Feature branch off `main` (e.g., `poc/word-tools`, `feat/<topic>`), PR back to `main`, squash merge.
- Conventional Commits: `feat:`, `fix:`, `chore:`, `test:`, `docs:`.
- Don't push to `main` directly.
