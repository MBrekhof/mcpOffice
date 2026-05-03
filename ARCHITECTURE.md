# Architecture

A landing-page map of mcpOffice. Stable content only — counts, status, and task lists live in `SESSION_HANDOFF.md` and `TODO.md`. Per-feature design lives in `docs/plans/`.

## What it is

A stdio MCP server (.NET 9, C#) that exposes Office-document operations as tools. Stateless: every tool call takes an absolute file path, performs one operation, returns. No session handles, no open-document lifecycle for the agent to manage.

## Layers

```
JSON-RPC over stdio
        |
        v
   Program.cs                 -- host bootstrap; AddMcpServer + WithToolsFromAssembly
        |
        v
   Tools/*Tools.cs            -- thin [McpServerTool] facade; one class per domain
        |
        v
   Services/<Domain>/         -- IXxxService + impl; the real work
        |                       (PathGuard, parse, transform, save)
        v
   DevExpress.Document.Processor
   (RichEditDocumentServer, SpreadsheetControl)
```

Cross-cutting:

- `Models/` — DTOs returned to the agent (records).
- `ErrorCode.cs` + `ToolError.cs` — stable string error codes encoded as `[code] message`.
- `PathGuard.cs` — absolute-path / exists / writable preconditions.

## Domains

| Domain  | Tool class            | Service                     | Tool prefix |
|---------|-----------------------|-----------------------------|-------------|
| Health  | `Tools/PingTools`     | —                           | (none)      |
| Word    | `Tools/WordTools`     | `Services/Word/WordDocumentService` | `word_`     |
| Excel   | `Tools/ExcelTools`    | `Services/Excel/ExcelWorkbookService` | `excel_`    |

The Excel VBA pipeline lives under `Services/Excel/Vba/` and has more moving parts than the rest of the codebase — see "VBA pipeline" below.

## Adding a new tool

The pattern, end-to-end (TDD):

1. **Decide domain.** Reuse existing `*Tools` + `*Service` if the tool fits a domain that already exists.
2. **Add DTOs** under `Models/` if the return shape is new. Records, file-scoped namespace, nullable enabled.
3. **Extend the service interface** (`IXxxService`) and impl. Service methods do path guarding, DevExpress work, and translate any non-`McpException` to `ToolError.ParseError(...)` (or a more specific code) inside a `try/catch (Exception ex) when (ex is not McpException)` wrapper.
4. **Write the failing unit test** under `tests/mcpOffice.Tests/<Domain>/`. Generate fixtures programmatically — see `Word/TestWordDocuments.cs` and `Excel/TestExcelWorkbooks.cs`. Don't commit `.docx`/`.xlsm` blobs unless there's a reason a programmatic fixture won't do.
5. **Implement** until the test passes.
6. **Add the `[McpServerTool(Name = "<domain>_<verb>")]`** method on the `*Tools` class. Always set `Name` explicitly (SDK lowercases otherwise). Keep tool methods as one-liners that delegate to the service.
7. **Update `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`** with the new tool name. This is the canary that catches accidental tool removal/rename.
8. **Conventional Commit**, feature branch off `main`, PR, squash-merge.

## Error model

Tools throw `McpException` only. The SDK doesn't expose a structured `Code` property in 1.2.0, so codes are encoded as a `[code]` prefix in the message:

```
[file_not_found] C:\path\to\missing.docx
```

Tests pattern-match on the prefix. Codes are stable identifiers an agent can branch on; the rest of the message is human-readable and may change. Codes live in `ErrorCode.cs`; helpers that throw them live in `ToolError.cs`.

## Stdio discipline

stdout is the JSON-RPC channel — anything written to it that isn't a valid frame breaks the client. Logging goes to **stderr only** via Serilog (configured in `Program.cs`). Never `Console.WriteLine` from tool or service code.

## Tests

Two suites:

- **`tests/mcpOffice.Tests/`** — unit tests against services directly. No transport. Fast. The bulk of correctness coverage lives here.
- **`tests/mcpOffice.Tests.Integration/`** — spawns the built server exe via stdio and round-trips through `McpClient`. Kept tiny on purpose (each test boots a process). Covers the tool catalog (`ToolSurfaceTests`), one happy path per workflow (`WordWorkflowTests`, `ExcelWorkflowTests`), and the ping smoke test.

Fixtures: programmatic generators in `TestWordDocuments.cs` / `TestExcelWorkbooks.cs`. Hand-authored binaries land in `tests/fixtures/` only when programmatic generation isn't viable (e.g. macro-enabled `.xlsm` with a real VBA project — DevExpress can't author one).

A gated "real-world" benchmark exists for VBA analysis: `Excel/Vba/AirSampleAnalysisTests.cs` runs against `C:\Projects\mcpOffice-samples\Air.xlsm` and skips when the file is absent.

## VBA pipeline (Excel)

`excel_extract_vba` and `excel_analyze_vba` share a layered pipeline under `Services/Excel/Vba/`:

```
.xlsm  -- (Open XML, DevExpress) -->  vbaProject.bin
                                          |
                                          v
                              MsOvbaDecompressor  (MS-OVBA RLE)
                                          |
                                          v
                              VbaDirStreamParser  (module records)
                                          |
                                          v
                              VbaProjectReader    (module sources, cp1252)
                                          |
                                          +--->  excel_extract_vba returns here
                                          |
                                          v
                              VbaSourceAnalyzer   (per-module orchestrator)
                                |        |        |
                                v        v        v
                              Cleaner  Scanner  CallGraph + ReferenceCollector
                                          |
                                          v
                              excel_analyze_vba returns the structural model
```

Strategy is regex-on-cleaned-source rather than a full VBA tokenizer. The Air.xlsm benchmark (107 modules, 938 call edges, 3040 object-model sites, ~115ms) is the evidence that this is sufficient. Revisit only if real-world ambiguity defeats the regex layer.

## DevExpress

- Package: `DevExpress.Document.Processor` (server-side, no UI). `RichEditDocumentServer` for Word, `SpreadsheetControl`/Open XML walks for Excel.
- Runtime license: `DevExpress_License.txt` at the repo root, gitignored.
- NuGet: nuget.org + a local filesystem source at `C:\Program Files\DevExpress 25.2\...\packages` (key `DevExpressLocal` in `nuget.config`). No URL feed with a token — that prompts for credentials in VS.

## What this architecture deliberately does not do

- No CI workflow yet — manual `dotnet build` / `dotnet test`.
- No multi-platform builds — Windows-only because of DevExpress runtime.
- No installer / signing.
- No retry on file locks — surfaces `io_error`, agent retries.
- No auth / sandboxing — runs locally as the user, full filesystem access.

## Where to look for what

| You want to know...                           | Look at                                              |
|-----------------------------------------------|------------------------------------------------------|
| Current branch state, last session            | `SESSION_HANDOFF.md`                                 |
| What's pending                                | `TODO.md`                                            |
| Why a feature is shaped the way it is         | `docs/plans/<date>-<feature>-design.md`              |
| How a feature was built                       | `docs/plans/<date>-<feature>-plan.md`                |
| How to wire the server into Claude Code       | `docs/usage.md`                                      |
| Project conventions, MCP SDK quirks, stdio    | `CLAUDE.md`                                          |
| The shape of the codebase (this file)         | `ARCHITECTURE.md`                                    |
