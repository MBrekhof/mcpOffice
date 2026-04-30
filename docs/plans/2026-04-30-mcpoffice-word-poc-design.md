# mcpOffice — Word POC Design

**Date:** 2026-04-30
**Status:** Approved (brainstorming phase)
**Scope:** Proof-of-concept covering Microsoft Word (.docx). Excel, PowerPoint, and PDF follow once the Word surface is proven.

## Purpose

An MCP server that lets AI agents manipulate Office documents through tool calls instead of writing throwaway Python scripts. Target consumer: Claude Code, Claude Desktop, and other MCP-capable agents running locally on Windows.

## Architecture

A .NET 9 console application using the official **`ModelContextProtocol`** C# SDK over **stdio** transport. Tools are static methods decorated with `[McpServerTool]`, auto-discovered via `WithToolsFromAssembly()`. Document operations sit behind an `IWordDocumentService` so the same surface can later be reused (or composed) for Excel / PowerPoint / PDF without touching tool definitions.

```
mcpOffice (console, stdio)
 |-- Program.cs            -- host bootstrap, MCP server registration
 |-- Tools/WordTools.cs    -- [McpServerTool] static methods (the API surface)
 |-- Services/Word/        -- IWordDocumentService + DevExpress-backed impl
 |     |-- WordReader.cs       -- Markdown + structured JSON projection
 |     |-- WordWriter.cs       -- create/modify
 |     `-- WordConverter.cs    -- docx <-> pdf / html / rtf / md
 |-- Models/               -- DTOs returned to the agent (records)
 `-- mcpOffice.csproj      -- refs DevExpress.Docs + ModelContextProtocol
```

DevExpress dependency: `DevExpress.Docs` (server-side, no UI) covers `RichEditDocumentServer` for Word plus the conversion pipeline. Assumes a DevExpress license is already installed on the machine.

## Operation model

**Stateless / file-path based.** Every tool call takes an absolute file path, performs one operation, writes back, and returns the path. No session handles, no open-document lifecycle for the agent to manage, no state to leak. The performance cost (open + parse on every call) is acceptable at agent speeds.

## Output format

Two read tools so agents can pick the right shape for the task:

- **Markdown** for "what does this say" queries — cheap, lossy, agent-native.
- **Structured JSON** for surgical edits — section/paragraph/run tree, tables, image refs, document properties.

## Tool surface (Word POC)

All paths absolute. All write/convert ops return the output path so calls can be chained. Booleans default to safe values (no overwrites, no destructive defaults).

### Reading

| Tool | Returns |
|---|---|
| `word_read_markdown(path)` | string — headings, lists, tables, emphasis, hyperlinks |
| `word_read_structured(path)` | JSON — `{ sections, tables, images, properties }` |
| `word_get_metadata(path)` | `{ author, title, subject, keywords, created, modified, lastPrinted, revisionCount, pageCount, wordCount }` |
| `word_list_comments(path)` | `[{ id, author, date, text, anchorText }]` |
| `word_list_revisions(path)` | tracked-changes summary |
| `word_get_outline(path)` | headings tree only (cheap) |

### Writing / creating

| Tool | Returns |
|---|---|
| `word_create_from_markdown(path, markdown, overwrite=false)` | path |
| `word_create_blank(path, overwrite=false)` | path |
| `word_append_markdown(path, markdown)` | path |
| `word_find_replace(path, find, replace, useRegex=false, matchCase=false)` | `{ replacements: int }` |
| `word_insert_paragraph(path, atIndex, text, style?)` | path |
| `word_insert_table(path, atIndex, headers[], rows[][])` | path |
| `word_set_metadata(path, properties{})` | path |
| `word_mail_merge(templatePath, outputPath, dataJson)` | path — tokens like `{{firstName}}` |

### Converting

| Tool | Returns |
|---|---|
| `word_convert(inputPath, outputPath, format)` | path — `format in { pdf, html, rtf, txt, markdown, docx }`, inferred from extension if omitted |

~15 tools. Composable, narrow, single-purpose. Heavy multi-step workflows are the agent's responsibility.

## Error model

Tools throw `McpException` with stable string error codes the agent can branch on, plus a human-readable message. The C# SDK serializes these as JSON-RPC errors.

| Code | When | Recovery hint |
|---|---|---|
| `file_not_found` | input path missing | re-check path / list directory |
| `file_exists` | output path exists, `overwrite=false` | retry with `overwrite=true` or different path |
| `invalid_path` | non-absolute, illegal chars | use absolute path |
| `unsupported_format` | extension/format not recognized | use one of `{pdf,html,rtf,txt,md,docx}` |
| `parse_error` | DevExpress can't load the doc (corrupt/locked) | message includes underlying reason |
| `index_out_of_range` | `atIndex` past end of doc | call `word_get_outline` first |
| `merge_field_missing` | template token has no value in `dataJson` | message names the field(s) |
| `io_error` | disk/permission failure | message includes OS error |
| `internal_error` | catch-all | bug — surface stack to logs only |

**Logging.** Stderr only (stdout is the JSON-RPC channel). Serilog console sink at `Information` by default; `Debug` for tool-call payloads when `MCPOFFICE_LOG=debug` env var is set. No file logging.

**License failures.** DevExpress license errors are caught at startup and surfaced as a clear stderr message ("DevExpress license not detected — install DevExpress.Docs and ensure licenses.licx is built in"). The server still starts so the client gets a useful tool-call error rather than a process that fails to launch.

## Testing

Two layers, both runnable from `dotnet test`.

### Unit / service layer (`mcpOffice.Tests`, xUnit + FluentAssertions)

- Targets `Services/Word/*` directly — no MCP transport involved.
- **Round-trip tests:** markdown -> docx -> markdown, structured read -> write -> structured read, asserting structural equivalence (not byte equality — DevExpress rewrites whitespace/IDs).
- **Fixtures** in `tests/fixtures/`: headings-only, table-heavy, with-comments, with-tracked-changes, with-merge-fields.
- **Converter goldens:** one per format (pdf, html, rtf, txt, md) — assert non-empty + magic-bytes / sniff for output type.
- **Error-path tests:** each error code has a test that triggers it and asserts the `McpException.Code`.

### Integration / protocol layer (`mcpOffice.Tests.Integration`)

- Spawns the built `mcpOffice.exe` as a child process, talks JSON-RPC over stdio using the C# SDK's `McpClient`.
- One end-to-end test per tool: client -> server -> assert response shape.
- One test that lists tools and verifies count + names match the source-of-truth list (catches accidental tool removal).

### Out of scope

- DevExpress's own correctness.
- Performance benchmarks.
- The MCP transport itself (covered by the SDK).

## Packaging & distribution

### Project layout

```
mcpOffice/
 |-- src/mcpOffice/                 -- the server (console app, net9.0)
 |-- tests/mcpOffice.Tests/         -- unit
 |-- tests/mcpOffice.Tests.Integration/
 |-- tests/fixtures/                -- sample .docx
 |-- docs/plans/                    -- design + implementation plan
 |-- docs/usage.md                  -- how to wire into Claude Code / Desktop
 |-- mcpOffice.sln
 |-- .gitignore                     -- bin/obj, *.user, licenses.licx variants
 `-- README.md
```

### NuGet dependencies (server)

- `ModelContextProtocol` — official C# MCP SDK (latest preview)
- `Microsoft.Extensions.Hosting` — for the host builder the SDK uses
- `DevExpress.Docs` — Word/Spreadsheet/PDF server-side APIs
- `Serilog.Extensions.Logging` + `Serilog.Sinks.Console` — stderr structured logs

DevExpress feed (`https://nuget.devexpress.com/<key>/api`) goes in a `nuget.config` at repo root; the auth key is the user's existing DevExpress license. Not committed.

### Local invocation

- `dotnet run --project src/mcpOffice` — development.
- `dotnet publish -c Release -r win-x64 --self-contained false` -> produces `mcpOffice.exe` for client config.

### Claude Code wiring (example, in `docs/usage.md`)

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\Projects\\mcpOffice\\src\\mcpOffice\\bin\\Release\\net9.0\\mcpOffice.exe"
    }
  }
}
```

### Out of scope for the POC

Installer, NuGet publishing, code signing, multi-platform builds. Revisit when Excel / PowerPoint / PDF are added.

## Open questions deferred to implementation

- Exact `RichEditDocumentServer` API for tracked-changes enumeration — verify during implementation.
- Markdown -> docx fidelity: DevExpress' built-in markdown import is recent and may need a small adapter for tables/strikethrough. Spike during the writer task.
- Whether `licenses.licx` needs to be in the test projects too, or only the server project. Confirm during first test run.
