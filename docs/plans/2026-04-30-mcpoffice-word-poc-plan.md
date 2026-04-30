# mcpOffice Word POC Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Ship a stdio MCP server in C# (.NET 9) that exposes ~15 tools for reading, writing, and converting Microsoft Word (.docx) documents using DevExpress.Docs, suitable for AI-agent consumption.

**Architecture:** Console app hosting `ModelContextProtocol` C# SDK over stdio. Static `[McpServerTool]` methods on `WordTools` delegate to an `IWordDocumentService` backed by `RichEditDocumentServer`. Stateless: every tool call takes an absolute file path. Errors raised as `McpException` with stable string codes.

**Tech Stack:** .NET 9 · ModelContextProtocol C# SDK · DevExpress.Docs · Serilog (stderr) · xUnit + FluentAssertions.

**Reference design:** `docs/plans/2026-04-30-mcpoffice-word-poc-design.md` — single source of truth for tool surface, error codes, and out-of-scope items. Read it before starting.

---

## Conventions used in this plan

- All paths are relative to `C:\Projects\mcpOffice\` (the repo root).
- Each tool follows the same TDD cycle defined once in **Task 8** ("Tool implementation pattern"). Subsequent tool tasks list only the fixture, signature, expected JSON shape, and notable error paths — they reuse that pattern verbatim.
- "Run tests" means: from repo root, `dotnet test --nologo --logger "console;verbosity=normal"`.
- Commits use Conventional Commits (`feat:`, `test:`, `chore:`, `docs:`).
- After every task: `dotnet build` is green AND every test passes. If either fails, stop and fix before moving on (per superpowers:verification-before-completion).

---

# Phase 0 — Repository bootstrap

### Task 1: Initialize git repo and baseline files

**Files:**
- Create: `.gitignore`
- Create: `README.md`
- Create: `nuget.config`

**Step 1: Init git**
```bash
cd C:/Projects/mcpOffice
git init
git branch -m main
```

**Step 2: Write `.gitignore`** (Visual Studio + .NET defaults plus DevExpress license artifacts)
```gitignore
bin/
obj/
*.user
*.suo
.vs/
*.licenses
licenses.licx.bak
*.log
TestResults/
```

**Step 3: Write minimal `README.md`**
```markdown
# mcpOffice

MCP server for Microsoft Office documents, backed by DevExpress.Docs.
Word (.docx) POC — see `docs/plans/2026-04-30-mcpoffice-word-poc-design.md`.
```

**Step 4: Write `nuget.config`**

The DevExpress feed key is the user's existing license; reference `%DXNUGET_KEY%` env var so it's never committed.
```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
    <add key="DevExpress" value="https://nuget.devexpress.com/%DXNUGET_KEY%/api" />
  </packageSources>
</configuration>
```

**Step 5: Commit**
```bash
git add .gitignore README.md nuget.config docs/
git commit -m "chore: scaffold repo with gitignore, readme, nuget config, design doc"
```

---

### Task 2: Create solution and three projects

**Files:**
- Create: `mcpOffice.sln`
- Create: `src/mcpOffice/mcpOffice.csproj`
- Create: `tests/mcpOffice.Tests/mcpOffice.Tests.csproj`
- Create: `tests/mcpOffice.Tests.Integration/mcpOffice.Tests.Integration.csproj`

**Step 1: Create projects via dotnet CLI**
```bash
dotnet new sln -n mcpOffice
dotnet new console -n mcpOffice -o src/mcpOffice -f net9.0
dotnet new xunit -n mcpOffice.Tests -o tests/mcpOffice.Tests -f net9.0
dotnet new xunit -n mcpOffice.Tests.Integration -o tests/mcpOffice.Tests.Integration -f net9.0
dotnet sln add src/mcpOffice/mcpOffice.csproj tests/mcpOffice.Tests/mcpOffice.Tests.csproj tests/mcpOffice.Tests.Integration/mcpOffice.Tests.Integration.csproj
```

**Step 2: Wire test projects to server project**
```bash
dotnet add tests/mcpOffice.Tests reference src/mcpOffice
dotnet add tests/mcpOffice.Tests.Integration reference src/mcpOffice
```

**Step 3: Verify build**
```bash
dotnet build --nologo
```
Expected: 3 projects build, 0 warnings, 0 errors.

**Step 4: Verify tests run** (xunit ships with one default `UnitTest1` — should pass)
```bash
dotnet test --nologo
```
Expected: `Passed: 2` (one default test in each test project).

**Step 5: Delete the default `UnitTest1.cs` files**
```bash
rm tests/mcpOffice.Tests/UnitTest1.cs tests/mcpOffice.Tests.Integration/UnitTest1.cs
```

**Step 6: Commit**
```bash
git add mcpOffice.sln src/ tests/
git commit -m "chore: add solution with server + unit + integration test projects"
```

---

### Task 3: Add NuGet package references

**Files:**
- Modify: `src/mcpOffice/mcpOffice.csproj`
- Modify: `tests/mcpOffice.Tests/mcpOffice.Tests.csproj`
- Modify: `tests/mcpOffice.Tests.Integration/mcpOffice.Tests.Integration.csproj`

**Step 1: Server packages**
```bash
dotnet add src/mcpOffice package ModelContextProtocol --prerelease
dotnet add src/mcpOffice package Microsoft.Extensions.Hosting
dotnet add src/mcpOffice package Serilog.Extensions.Hosting
dotnet add src/mcpOffice package Serilog.Sinks.Console
dotnet add src/mcpOffice package DevExpress.Docs
```

**Step 2: Test packages**
```bash
dotnet add tests/mcpOffice.Tests package FluentAssertions
dotnet add tests/mcpOffice.Tests package DevExpress.Docs
dotnet add tests/mcpOffice.Tests.Integration package FluentAssertions
dotnet add tests/mcpOffice.Tests.Integration package ModelContextProtocol --prerelease
```

**Step 3: Build to confirm restore works**
```bash
dotnet build --nologo
```
Expected: 0 warnings, 0 errors. If DevExpress restore fails, the user's `DXNUGET_KEY` env var is missing or wrong — stop and surface this.

**Step 4: Commit**
```bash
git add src/mcpOffice/mcpOffice.csproj tests/**/*.csproj
git commit -m "chore: add MCP SDK, DevExpress.Docs, Serilog, FluentAssertions"
```

---

# Phase 1 — MCP server bootstrap

### Task 4: Minimal Program.cs with stdio MCP host

**Files:**
- Modify: `src/mcpOffice/Program.cs`
- Create: `src/mcpOffice/Tools/PingTools.cs`

**Step 1: Replace `Program.cs`**
```csharp
using McpOffice.Tools;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Information()
    .WriteTo.Console(standardErrorFromLevel: Serilog.Events.LogEventLevel.Verbose)
    .CreateLogger();

var builder = Host.CreateApplicationBuilder(args);
builder.Logging.ClearProviders();
builder.Logging.AddSerilog();

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync();
```

**Step 2: Add a single `ping` tool to prove wiring**
```csharp
// src/mcpOffice/Tools/PingTools.cs
using System.ComponentModel;
using ModelContextProtocol.Server;

namespace McpOffice.Tools;

[McpServerToolType]
public static class PingTools
{
    [McpServerTool, Description("Returns 'pong'. Use to verify the server is reachable.")]
    public static string Ping() => "pong";
}
```

**Step 3: Build**
```bash
dotnet build --nologo
```
Expected: success.

**Step 4: Smoke-run the server** (verifies stdio loop starts)
```bash
echo "" | dotnet run --project src/mcpOffice --no-build
```
Expected: process starts, blocks on stdin (empty echo gives EOF -> graceful exit). No exceptions.

**Step 5: Commit**
```bash
git add src/mcpOffice/Program.cs src/mcpOffice/Tools/PingTools.cs
git commit -m "feat: stdio MCP host with ping tool"
```

---

### Task 5: Integration test that spawns server and calls ping

**Files:**
- Create: `tests/mcpOffice.Tests.Integration/ServerHarness.cs`
- Create: `tests/mcpOffice.Tests.Integration/PingTests.cs`

**Step 1: Write `ServerHarness.cs`** — spawns the published server exe and gives back an `IMcpClient`. Uses `StdioClientTransport` from the SDK.

```csharp
using ModelContextProtocol.Client;

namespace McpOffice.Tests.Integration;

public sealed class ServerHarness : IAsyncDisposable
{
    public IMcpClient Client { get; private set; } = null!;

    public static async Task<ServerHarness> StartAsync()
    {
        var serverExe = ResolveServerDll();
        var transport = new StdioClientTransport(new()
        {
            Name = "mcpOffice",
            Command = "dotnet",
            Arguments = [serverExe]
        });
        var harness = new ServerHarness();
        harness.Client = await McpClientFactory.CreateAsync(transport);
        return harness;
    }

    private static string ResolveServerDll()
    {
        // Walk up from the test bin dir to repo root, then into the server bin.
        var asmDir = Path.GetDirectoryName(typeof(ServerHarness).Assembly.Location)!;
        var repoRoot = new DirectoryInfo(asmDir);
        while (repoRoot is not null && !File.Exists(Path.Combine(repoRoot.FullName, "mcpOffice.sln")))
            repoRoot = repoRoot.Parent;
        if (repoRoot is null) throw new InvalidOperationException("Could not locate repo root.");
        var dll = Path.Combine(repoRoot.FullName, "src", "mcpOffice", "bin", "Debug", "net9.0", "mcpOffice.dll");
        if (!File.Exists(dll)) throw new FileNotFoundException($"Server build output missing: {dll}. Run 'dotnet build' first.");
        return dll;
    }

    public async ValueTask DisposeAsync() => await Client.DisposeAsync();
}
```

**Step 2: Write the failing test**
```csharp
// tests/mcpOffice.Tests.Integration/PingTests.cs
using FluentAssertions;
using ModelContextProtocol.Client;

namespace McpOffice.Tests.Integration;

public class PingTests
{
    [Fact]
    public async Task Ping_returns_pong()
    {
        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync("Ping", new Dictionary<string, object?>());
        var text = result.Content.OfType<ModelContextProtocol.Protocol.TextContentBlock>().Single().Text;
        text.Should().Be("pong");
    }

    [Fact]
    public async Task Lists_ping_tool()
    {
        await using var harness = await ServerHarness.StartAsync();
        var tools = await harness.Client.ListToolsAsync();
        tools.Should().Contain(t => t.Name == "Ping");
    }
}
```

**Step 3: Run — should pass**
```bash
dotnet build --nologo && dotnet test tests/mcpOffice.Tests.Integration --nologo
```
Expected: `Passed: 2`. If transport/IPC fails, surface the stderr — likely missing build output.

**Step 4: Commit**
```bash
git add tests/mcpOffice.Tests.Integration/
git commit -m "test: integration harness + ping round-trip via stdio"
```

---

# Phase 2 — Word service foundations

### Task 6: Error model — McpToolException with stable codes

**Files:**
- Create: `src/mcpOffice/ErrorCode.cs`
- Create: `src/mcpOffice/ToolError.cs`
- Create: `tests/mcpOffice.Tests/ToolErrorTests.cs`

**Step 1: Write the failing test**
```csharp
// tests/mcpOffice.Tests/ToolErrorTests.cs
using FluentAssertions;
using McpOffice;
using ModelContextProtocol;

namespace McpOffice.Tests;

public class ToolErrorTests
{
    [Fact]
    public void FileNotFound_throws_McpException_with_code_in_message()
    {
        var act = () => ToolError.FileNotFound("C:\\missing.docx");
        act.Should().Throw<McpException>()
           .Which.Message.Should().Contain("file_not_found").And.Contain("C:\\missing.docx");
    }
}
```

**Step 2: Run — fails (types don't exist)**
```bash
dotnet test tests/mcpOffice.Tests --nologo
```

**Step 3: Implement**
```csharp
// src/mcpOffice/ErrorCode.cs
namespace McpOffice;

public static class ErrorCode
{
    public const string FileNotFound = "file_not_found";
    public const string FileExists = "file_exists";
    public const string InvalidPath = "invalid_path";
    public const string UnsupportedFormat = "unsupported_format";
    public const string ParseError = "parse_error";
    public const string IndexOutOfRange = "index_out_of_range";
    public const string MergeFieldMissing = "merge_field_missing";
    public const string IoError = "io_error";
    public const string InternalError = "internal_error";
}
```

```csharp
// src/mcpOffice/ToolError.cs
using ModelContextProtocol;

namespace McpOffice;

public static class ToolError
{
    public static Exception FileNotFound(string path) =>
        Throw(ErrorCode.FileNotFound, $"File not found: {path}");
    public static Exception FileExists(string path) =>
        Throw(ErrorCode.FileExists, $"Output already exists (pass overwrite=true to replace): {path}");
    public static Exception InvalidPath(string path) =>
        Throw(ErrorCode.InvalidPath, $"Path must be absolute and well-formed: {path}");
    public static Exception UnsupportedFormat(string format) =>
        Throw(ErrorCode.UnsupportedFormat, $"Unsupported format: {format}. Use one of pdf, html, rtf, txt, markdown, docx.");
    public static Exception ParseError(string path, string detail) =>
        Throw(ErrorCode.ParseError, $"Could not parse {path}: {detail}");
    public static Exception IndexOutOfRange(int index, int max) =>
        Throw(ErrorCode.IndexOutOfRange, $"Index {index} is out of range (0..{max}).");
    public static Exception MergeFieldMissing(IEnumerable<string> fields) =>
        Throw(ErrorCode.MergeFieldMissing, $"Template fields with no value in dataJson: {string.Join(", ", fields)}");
    public static Exception IoError(string detail) =>
        Throw(ErrorCode.IoError, $"IO error: {detail}");
    public static Exception Internal(string detail) =>
        Throw(ErrorCode.InternalError, $"Internal error: {detail}");

    private static McpException Throw(string code, string message) =>
        new($"[{code}] {message}");
}
```

> Note: the C# SDK's `McpException` doesn't expose a structured `Code` property today, so we encode the code as a `[code]` prefix in the message. Agents and tests can pattern-match on it. Revisit if the SDK gains a typed code field.

**Step 4: Run — passes**
```bash
dotnet test tests/mcpOffice.Tests --nologo
```

**Step 5: Commit**
```bash
git add src/mcpOffice/ErrorCode.cs src/mcpOffice/ToolError.cs tests/mcpOffice.Tests/ToolErrorTests.cs
git commit -m "feat: stable error codes via ToolError helper"
```

---

### Task 7: PathGuard — absolute-path / file-existence checks

**Files:**
- Create: `src/mcpOffice/PathGuard.cs`
- Create: `tests/mcpOffice.Tests/PathGuardTests.cs`

**Step 1: Write failing tests**
```csharp
// tests/mcpOffice.Tests/PathGuardTests.cs
using FluentAssertions;
using McpOffice;
using ModelContextProtocol;

namespace McpOffice.Tests;

public class PathGuardTests
{
    [Fact]
    public void RequireAbsolute_rejects_relative()
    {
        var act = () => PathGuard.RequireAbsolute("foo.docx");
        act.Should().Throw<McpException>().Which.Message.Should().Contain("invalid_path");
    }

    [Fact]
    public void RequireAbsolute_accepts_absolute()
    {
        PathGuard.RequireAbsolute("C:\\foo.docx"); // does not throw
    }

    [Fact]
    public void RequireExists_throws_when_missing()
    {
        var act = () => PathGuard.RequireExists("C:\\definitely-does-not-exist-xyz.docx");
        act.Should().Throw<McpException>().Which.Message.Should().Contain("file_not_found");
    }

    [Fact]
    public void RequireWritable_throws_when_exists_and_no_overwrite()
    {
        var tmp = Path.Combine(Path.GetTempPath(), $"pg-{Guid.NewGuid():N}.tmp");
        File.WriteAllText(tmp, "x");
        try
        {
            var act = () => PathGuard.RequireWritable(tmp, overwrite: false);
            act.Should().Throw<McpException>().Which.Message.Should().Contain("file_exists");
        }
        finally { File.Delete(tmp); }
    }
}
```

**Step 2: Implement**
```csharp
// src/mcpOffice/PathGuard.cs
namespace McpOffice;

public static class PathGuard
{
    public static void RequireAbsolute(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path))
            throw ToolError.InvalidPath(path);
    }

    public static void RequireExists(string path)
    {
        RequireAbsolute(path);
        if (!File.Exists(path)) throw ToolError.FileNotFound(path);
    }

    public static void RequireWritable(string path, bool overwrite)
    {
        RequireAbsolute(path);
        if (File.Exists(path) && !overwrite) throw ToolError.FileExists(path);
        var dir = Path.GetDirectoryName(path);
        if (dir is not null) Directory.CreateDirectory(dir);
    }
}
```

**Step 3: Run — passes**
```bash
dotnet test tests/mcpOffice.Tests --nologo
```

**Step 4: Commit**
```bash
git add src/mcpOffice/PathGuard.cs tests/mcpOffice.Tests/PathGuardTests.cs
git commit -m "feat: PathGuard for absolute/exists/writable preconditions"
```

---

### Task 8: Tool implementation pattern (template — read carefully)

This task implements the **first** Word tool (`word_get_outline`) end-to-end. Every subsequent tool task follows the exact same five steps; later tasks will list only the deltas (test fixture, signature, expected JSON shape, error paths).

**Files:**
- Create: `src/mcpOffice/Services/Word/IWordDocumentService.cs`
- Create: `src/mcpOffice/Services/Word/WordDocumentService.cs`
- Create: `src/mcpOffice/Models/OutlineNode.cs`
- Create: `src/mcpOffice/Tools/WordTools.cs`
- Create: `tests/mcpOffice.Tests/Word/OutlineTests.cs`
- Create: `tests/fixtures/headings-only.docx` (generated below)

**Step 1: Write failing test**

```csharp
// tests/mcpOffice.Tests/Word/OutlineTests.cs
using FluentAssertions;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class OutlineTests
{
    [Fact]
    public void Outline_returns_heading_tree()
    {
        var svc = new WordDocumentService();
        var fixture = TestFixtures.Path("headings-only.docx");
        var nodes = svc.GetOutline(fixture);
        nodes.Should().NotBeEmpty();
        nodes[0].Level.Should().Be(1);
        nodes[0].Text.Should().NotBeNullOrWhiteSpace();
    }
}
```

`TestFixtures` is a tiny helper:
```csharp
// tests/mcpOffice.Tests/TestFixtures.cs
namespace McpOffice.Tests;
public static class TestFixtures
{
    public static string Path(string name)
    {
        var asmDir = System.IO.Path.GetDirectoryName(typeof(TestFixtures).Assembly.Location)!;
        var dir = new DirectoryInfo(asmDir);
        while (dir is not null && !File.Exists(System.IO.Path.Combine(dir.FullName, "mcpOffice.sln")))
            dir = dir.Parent;
        return System.IO.Path.Combine(dir!.FullName, "tests", "fixtures", name);
    }
}
```

**Step 2: Run — fails (no service yet)**
```bash
dotnet test tests/mcpOffice.Tests --nologo
```

**Step 3: Generate fixture** — write a one-shot helper test that creates `headings-only.docx`, run it once, commit the binary, then delete the helper (or keep it under `[Trait("category", "fixture-gen")]`).

```csharp
// scratch — run once, then remove or skip
[Fact(Skip = "fixture generator — unskip to regenerate")]
public void Generate_headings_only_fixture()
{
    using var srv = new DevExpress.XtraRichEdit.RichEditDocumentServer();
    var doc = srv.Document;
    doc.AppendText("Introduction"); doc.Paragraphs.Last().Style = doc.ParagraphStyles["Heading 1"];
    doc.AppendText("\nBackground"); doc.Paragraphs.Last().Style = doc.ParagraphStyles["Heading 2"];
    doc.AppendText("\nDetails"); doc.Paragraphs.Last().Style = doc.ParagraphStyles["Heading 3"];
    doc.AppendText("\nConclusion"); doc.Paragraphs.Last().Style = doc.ParagraphStyles["Heading 1"];
    srv.SaveDocument(TestFixtures.Path("headings-only.docx"), DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
}
```

**Step 4: Implement service + DTO + tool**

```csharp
// src/mcpOffice/Models/OutlineNode.cs
namespace McpOffice.Models;
public sealed record OutlineNode(int Level, string Text);
```

```csharp
// src/mcpOffice/Services/Word/IWordDocumentService.cs
using McpOffice.Models;
namespace McpOffice.Services.Word;

public interface IWordDocumentService
{
    IReadOnlyList<OutlineNode> GetOutline(string path);
    // ... (added incrementally per task)
}
```

```csharp
// src/mcpOffice/Services/Word/WordDocumentService.cs
using DevExpress.XtraRichEdit;
using McpOffice.Models;

namespace McpOffice.Services.Word;

public sealed class WordDocumentService : IWordDocumentService
{
    public IReadOnlyList<OutlineNode> GetOutline(string path)
    {
        PathGuard.RequireExists(path);
        try
        {
            using var srv = new RichEditDocumentServer();
            srv.LoadDocument(path);
            var doc = srv.Document;
            var outline = new List<OutlineNode>();
            foreach (var para in doc.Paragraphs)
            {
                var styleName = para.Style?.Name ?? "";
                if (styleName.StartsWith("Heading ") &&
                    int.TryParse(styleName["Heading ".Length..], out var level))
                {
                    outline.Add(new OutlineNode(level, doc.GetText(para.Range).Trim()));
                }
            }
            return outline;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }
}
```

```csharp
// src/mcpOffice/Tools/WordTools.cs
using System.ComponentModel;
using McpOffice.Services.Word;
using ModelContextProtocol.Server;

namespace McpOffice.Tools;

[McpServerToolType]
public static class WordTools
{
    private static readonly IWordDocumentService Svc = new WordDocumentService();

    [McpServerTool(Name = "word_get_outline"),
     Description("Returns the heading tree of a .docx file as [{level,text}]. Cheap; use to skim structure.")]
    public static object WordGetOutline(
        [Description("Absolute path to the .docx file")] string path)
        => Svc.GetOutline(path);
}
```

**Step 5: Run tests + commit**
```bash
dotnet test --nologo
git add src/mcpOffice/ tests/mcpOffice.Tests/Word/ tests/mcpOffice.Tests/TestFixtures.cs tests/fixtures/headings-only.docx
git commit -m "feat: word_get_outline + WordDocumentService skeleton"
```

**The pattern, restated:**
1. Add fixture (or reuse one). 2. Write unit test against `WordDocumentService`. 3. Run — fail. 4. Implement (extend `IWordDocumentService`, extend `WordDocumentService`, add `[McpServerTool]` on `WordTools`). 5. Run — pass. 6. Commit.

---

# Phase 3 — Read tools

> Each task below = one tool, following Task 8 pattern. Listed: fixture, service signature, tool signature, error paths.

### Task 9: `word_get_metadata`

- **Fixture:** reuse `headings-only.docx`; extend the generator to set `doc.DocumentProperties.Author = "Martin"`, `Title`, `Subject`, `Keywords`.
- **DTO:** `record DocumentMetadata(string? Author, string? Title, string? Subject, string? Keywords, DateTime? Created, DateTime? Modified, DateTime? LastPrinted, int RevisionCount, int PageCount, int WordCount)` — pull from `doc.DocumentProperties` and `srv.Document.GetText(srv.Document.Range).Split(...).Length` for word count; `srv.GetPageCount()` for pages.
- **Service:** `DocumentMetadata GetMetadata(string path)`.
- **Tool:** `word_get_metadata(path) -> DocumentMetadata`.
- **Error paths covered by Task 7's PathGuard:** `file_not_found`, `invalid_path`.

### Task 10: `word_read_markdown`

- **Fixture:** create `tests/fixtures/mixed.docx` with a heading, paragraph with bold/italic, bullet list, simple 2x2 table, hyperlink.
- **Service:** `string ReadAsMarkdown(string path)` — uses `srv.Options.Export.Markdown` and `srv.SaveDocument(stream, DocumentFormat.Markdown)` then `Encoding.UTF8.GetString`.
- **Tool:** `word_read_markdown(path) -> string`.
- **Test asserts:** output contains `# `, `**bold**`, `- ` bullets, table pipe markers.

### Task 11: `word_read_structured`

- **Fixture:** `mixed.docx`.
- **DTO tree:** `record StructuredDocument(IReadOnlyList<Block> Blocks, IReadOnlyList<TableBlock> Tables, IReadOnlyList<ImageRef> Images, DocumentMetadata Properties);` with `abstract record Block` -> `HeadingBlock(int Level, string Text)`, `ParagraphBlock(IReadOnlyList<Run> Runs)`, `Run(string Text, bool Bold, bool Italic, string? HyperlinkUrl)`, `TableBlock(int Index, IReadOnlyList<IReadOnlyList<string>> Rows)`, `ImageRef(int Index, string? AltText)`.
- **Service:** `StructuredDocument ReadStructured(string path)` — walks `doc.Paragraphs`, then `doc.Tables`, then `doc.Images`.
- **Tool:** `word_read_structured(path) -> StructuredDocument` (returned as JSON).
- **Test:** at least one heading, one paragraph with a bold run, one table with 2 rows.

### Task 12: `word_list_comments`

- **Fixture:** create `tests/fixtures/with-comments.docx` — two comments with different authors.
- **DTO:** `record CommentEntry(int Id, string Author, DateTime Date, string Text, string AnchorText)`.
- **Service:** `IReadOnlyList<CommentEntry> ListComments(string path)` — iterate `doc.Comments`.
- **Tool:** `word_list_comments(path)`.

### Task 13: `word_list_revisions`

- **Fixture:** `tests/fixtures/with-tracked-changes.docx` — generate by enabling `doc.RevisionOptions.ShowMarkup = true`, making an edit while `srv.Document.RevisionsEnabled = true`.
- **DTO:** `record RevisionEntry(string Type, string Author, DateTime Date, string Text)` (Type ∈ "insert","delete","format").
- **Service:** `IReadOnlyList<RevisionEntry> ListRevisions(string path)` — iterate `doc.Revisions`.
- **Tool:** `word_list_revisions(path)`.

---

# Phase 4 — Write / create tools

### Task 14: `word_create_blank`

- **Service:** `string CreateBlank(string path, bool overwrite)` — `PathGuard.RequireWritable(path, overwrite)`, instantiate empty `RichEditDocumentServer`, `SaveDocument(path, DocumentFormat.OpenXml)`, return path.
- **Tool:** `word_create_blank(path, overwrite=false)`.
- **Test:** create then assert file exists and is a valid .docx (`new RichEditDocumentServer().LoadDocument(path)` doesn't throw).
- **Error path test:** call twice without `overwrite=true` -> `file_exists`.

### Task 15: `word_create_from_markdown`

- **Service:** `string CreateFromMarkdown(string path, string markdown, bool overwrite)` — `srv.LoadDocument(Encoding.UTF8.GetBytes(markdown), DocumentFormat.Markdown)`, save as OpenXml.
- **Tool:** `word_create_from_markdown(path, markdown, overwrite=false)`.
- **Test:** round-trip — create from markdown `"# Title\n\nHello **world**"`, then `ReadAsMarkdown` and assert it contains `# Title` and `**world**`.

### Task 16: `word_append_markdown`

- **Service:** `string AppendMarkdown(string path, string markdown)` — load, `srv.Document.AppendDocumentContent(...)`, save back.
- **Tool:** `word_append_markdown(path, markdown)`.
- **Test:** create blank, append `"# H"`, read outline, expect 1 heading.

### Task 17: `word_find_replace`

- **DTO:** `record ReplaceResult(int Replacements)`.
- **Service:** `ReplaceResult FindReplace(string path, string find, string replace, bool useRegex, bool matchCase)` — use `doc.FindAll(...)` then `doc.Replace(...)`. Save in place.
- **Tool:** `word_find_replace(path, find, replace, useRegex=false, matchCase=false)`.
- **Test:** create from markdown `"hello hello"`, replace `"hello"` -> `"hi"`, expect `Replacements == 2` and content contains `"hi hi"`.

### Task 18: `word_insert_paragraph`

- **Service:** `string InsertParagraph(string path, int atIndex, string text, string? style)` — bounds-check `atIndex` against `doc.Paragraphs.Count`; throw `IndexOutOfRange` if past end.
- **Tool:** `word_insert_paragraph(path, atIndex, text, style?)`.
- **Test:** insert at 0 with style "Heading 1", outline grows by 1.
- **Error test:** `atIndex = 999` -> `index_out_of_range`.

### Task 19: `word_insert_table`

- **Service:** `string InsertTable(string path, int atIndex, IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows)` — `doc.Tables.Create(...)`, fill cells.
- **Tool:** `word_insert_table(path, atIndex, headers[], rows[][])`.
- **Test:** insert a 2x2 table, structured read returns one `TableBlock` with matching cells.

### Task 20: `word_set_metadata`

- **Service:** `string SetMetadata(string path, IReadOnlyDictionary<string,string> properties)` — accept keys: `author`, `title`, `subject`, `keywords`. Unknown keys -> `unsupported_format` (or new code `unknown_property` — discuss; keep `unsupported_format` for now).
- **Tool:** `word_set_metadata(path, properties)`.
- **Test:** set `author=Bob`, `GetMetadata` returns `Bob`.

### Task 21: `word_mail_merge`

- **Service:** `string MailMerge(string templatePath, string outputPath, string dataJson)` — parse JSON to `Dictionary<string,string>`, find tokens `{{name}}` in document text, replace each, save to output. Collect missing tokens; if any, throw `MergeFieldMissing`.
- **Tool:** `word_mail_merge(templatePath, outputPath, dataJson)`.
- **Test:** template `"Hello {{firstName}}!"`, data `{"firstName":"Ada"}`, output reads `"Hello Ada!"`.
- **Error test:** missing field -> `merge_field_missing`.

---

# Phase 5 — Convert

### Task 22: `word_convert`

- **Service:** `string Convert(string inputPath, string outputPath, string? format)` — if `format` null, infer from `Path.GetExtension(outputPath)`. Map: `.pdf` -> `srv.ExportToPdf(stream)`, `.html` -> `DocumentFormat.Html`, `.rtf` -> `DocumentFormat.Rtf`, `.txt` -> `DocumentFormat.PlainText`, `.md`/`.markdown` -> `DocumentFormat.Markdown`, `.docx` -> `DocumentFormat.OpenXml`.
- **Tool:** `word_convert(inputPath, outputPath, format?)`.
- **Tests (one per format):** convert `mixed.docx` -> each output, assert non-empty + magic bytes:
  - PDF: starts with `%PDF-`
  - HTML: contains `<html`
  - RTF: starts with `{\\rtf`
  - TXT: contains heading text, no markup
  - MD: contains `# `
  - DOCX: ZIP magic `PK\x03\x04`
- **Error test:** `format = "xyz"` -> `unsupported_format`.

---

# Phase 6 — Integration tests + docs

### Task 23: Tool-surface integration test

**File:** `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`

Asserts the tool catalog the server exposes matches the design doc exactly. Catches accidental tool removal/rename.

```csharp
[Fact]
public async Task Exposes_expected_tool_catalog()
{
    string[] expected = [
        "Ping",
        "word_get_outline", "word_get_metadata", "word_read_markdown",
        "word_read_structured", "word_list_comments", "word_list_revisions",
        "word_create_blank", "word_create_from_markdown", "word_append_markdown",
        "word_find_replace", "word_insert_paragraph", "word_insert_table",
        "word_set_metadata", "word_mail_merge", "word_convert"
    ];
    await using var harness = await ServerHarness.StartAsync();
    var tools = (await harness.Client.ListToolsAsync()).Select(t => t.Name).ToHashSet();
    tools.Should().BeEquivalentTo(expected);
}
```

### Task 24: One end-to-end integration test per tool group

Three tests — one read, one write, one convert — that prove the JSON-RPC layer doesn't drop anything. Don't re-test every tool through the protocol; the unit tests cover correctness, these cover transport.

- `Read_markdown_round_trip_via_stdio` — copy `mixed.docx` to temp, call `word_read_markdown`, assert response text contains `# `.
- `Create_then_outline_via_stdio` — call `word_create_from_markdown` then `word_get_outline` on the same temp path, expect 1 heading.
- `Convert_to_pdf_via_stdio` — call `word_convert` with `.pdf` output, read first 5 bytes of result file, expect `%PDF-`.

### Task 25: Docs — `docs/usage.md` and update README

**Files:**
- Create: `docs/usage.md` — install, build, Claude Code config snippet, env vars (`DXNUGET_KEY`, `MCPOFFICE_LOG`), troubleshooting (license errors, locked files).
- Modify: `README.md` — link to design doc, plan, usage doc; show one example tool call.

**Commit:**
```bash
git add docs/usage.md README.md
git commit -m "docs: usage guide + README polish"
```

### Task 26: Final verification

**Steps:**
1. `dotnet build -c Release --nologo` — clean build.
2. `dotnet test -c Release --nologo` — every test passes.
3. `dotnet publish -c Release -r win-x64 --self-contained false src/mcpOffice` — produces `mcpOffice.exe`.
4. Manually wire into Claude Code (`%APPDATA%\Claude\claude_desktop_config.json` or project-local `.mcp.json`) and call `word_get_outline` against a real .docx in this very session — verify it works end-to-end with a live agent. Per global CLAUDE.md: **build green != it works**; verify with a real agent call.
5. Update `MEMORY.md` and add a SESSION_HANDOFF.md noting POC complete + next milestone (Excel).

---

## What this plan deliberately does NOT do

- No CI workflow (Task: add later when Excel/PowerPoint land).
- No DevExpress license install automation — assumed pre-installed.
- No telemetry / metrics.
- No multi-platform builds.
- No installer / MSI / signing.
- No retry logic on file locks — first-call failure surfaces `io_error`; agent retries.

## Risks called out

1. **DevExpress markdown round-trip fidelity** — first time we hit `Markdown` format both ways. Spike during Task 10 / 15; if fidelity is poor for tables, fall back to a hand-rolled markdown writer for the create path.
2. **`McpException.Code` not yet typed** — current SDK encodes via message prefix. If/when the SDK adds a `Code` property, refactor `ToolError` (single file, low risk).
3. **Stdio integration tests are slow** — each spawns `dotnet`. Keep the integration suite tiny (Tasks 5, 23, 24) and rely on unit tests for breadth.
