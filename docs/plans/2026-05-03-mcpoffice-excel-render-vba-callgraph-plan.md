# excel_render_vba_callgraph Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship a new MCP tool `excel_render_vba_callgraph` that emits the VBA call graph as Mermaid or DOT for visual inspection, layered on the existing `excel_analyze_vba` analyzer. Reference design: `docs/plans/2026-05-03-mcpoffice-excel-render-vba-callgraph-design.md`.

**Architecture:** Pure-function filter (`VbaCallgraphFilter`) takes the analyzer's full `ExcelVbaAnalysis`, applies module / procedure / depth / direction filters, classifies orphans/externals, and emits a format-agnostic `FilteredCallgraph`. Two renderers (`MermaidCallgraphRenderer`, `DotCallgraphRenderer`) implement a common `ICallgraphRenderer` interface and serialise that DTO to a string. `excel_analyze_vba` is unchanged; this work is purely additive.

**Tech Stack:** .NET 9, C#, ModelContextProtocol C# SDK, xUnit. No new NuGet packages. Reuses existing analyzer pipeline under `src/mcpOffice/Services/Excel/Vba/`.

---

## Conventions

- All paths are relative to the repo root (`C:\Projects\mcpOffice\` on the dev machine).
- "Run tests" means `dotnet test --nologo --logger "console;verbosity=normal"` from repo root unless a specific test project is named.
- After each task: `dotnet build` and `dotnet test` must both be green before moving on.
- Conventional Commits: `feat:`, `test:`, `chore:`, `docs:`.
- Per the project CLAUDE.md, work happens on a feature branch (`feat/render-vba-callgraph`), squash-merged via PR. Don't commit on `main`.

---

## File structure

**New production files:**

| Path | Purpose |
|---|---|
| `src/mcpOffice/Models/CallgraphNode.cs` | Node DTO carried in `FilteredCallgraph` — id, label, module, classification flags |
| `src/mcpOffice/Models/CallgraphEdge.cs` | Edge DTO — fromId, toId, resolved |
| `src/mcpOffice/Models/FilteredCallgraph.cs` | Filter output carried into the renderers |
| `src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs` | Pure static filter: analysis + options → FilteredCallgraph |
| `src/mcpOffice/Services/Excel/Vba/Rendering/ICallgraphRenderer.cs` | Renderer interface |
| `src/mcpOffice/Services/Excel/Vba/Rendering/MermaidCallgraphRenderer.cs` | Mermaid impl |
| `src/mcpOffice/Services/Excel/Vba/Rendering/DotCallgraphRenderer.cs` | DOT impl |

**Modified production files:**

| Path | Change |
|---|---|
| `src/mcpOffice/ErrorCode.cs` | Add `ProcedureNotFound`, `GraphTooLarge`, `InvalidRenderOption` |
| `src/mcpOffice/ToolError.cs` | Add helper throws for the new codes |
| `src/mcpOffice/Services/Excel/IExcelWorkbookService.cs` | Add `RenderVbaCallgraph(...)` |
| `src/mcpOffice/Services/Excel/ExcelWorkbookService.cs` | Implement `RenderVbaCallgraph` |
| `src/mcpOffice/Tools/ExcelTools.cs` | Add `[McpServerTool(Name="excel_render_vba_callgraph")]` method |

**New test files:**

| Path | Purpose |
|---|---|
| `tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs` | Filter unit tests against synthetic `ExcelVbaAnalysis` |
| `tests/mcpOffice.Tests/Excel/Vba/Rendering/MermaidCallgraphRendererTests.cs` | Mermaid renderer unit tests |
| `tests/mcpOffice.Tests/Excel/Vba/Rendering/DotCallgraphRendererTests.cs` | DOT renderer unit tests |
| `tests/mcpOffice.Tests/Excel/Vba/AirSampleRenderTests.cs` | Gated benchmark against `C:\Projects\mcpOffice-samples\Air.xlsm` |

**Modified test files:**

| Path | Change |
|---|---|
| `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs` | Add stdio test for `excel_render_vba_callgraph` |
| `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs` | Include `excel_render_vba_callgraph` in expected catalog (24 → 25 tools) |

---

# Phase 0 — Branch

### Task 1: Create the feature branch

**Files:** —

- [ ] **Step 1: Branch off main**

```bash
git checkout main
git pull --ff-only
git checkout -b feat/render-vba-callgraph
```

- [ ] **Step 2: Confirm clean state**

```bash
git status
dotnet build --nologo
dotnet test --nologo
```

Expected: `nothing to commit, working tree clean`. Build green. Tests green (145 passing).

---

# Phase 1 — Error codes

### Task 2: Add three new error codes and helpers

**Files:**
- Modify: `src/mcpOffice/ErrorCode.cs`
- Modify: `src/mcpOffice/ToolError.cs`
- Test: `tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs` (file already exists; add three tests)

- [ ] **Step 1: Read the existing error code test file**

```
Read tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs
```

Note the pattern — each test calls a `ToolError.Xxx(...)` helper and asserts the message contains `[code]` plus context.

- [ ] **Step 2: Write three failing tests**

Append to `tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs`:

```csharp
[Fact]
public void ProcedureNotFound_throws_McpException_with_code_and_candidates()
{
    var act = () => ToolError.ProcedureNotFound("ReadExports", new[] { "SaveDB", "Paste2Cell" });
    act.Should().Throw<ModelContextProtocol.McpException>()
       .Which.Message.Should().Contain("procedure_not_found")
       .And.Contain("ReadExports")
       .And.Contain("SaveDB")
       .And.Contain("Paste2Cell");
}

[Fact]
public void GraphTooLarge_throws_McpException_with_count_and_max()
{
    var act = () => ToolError.GraphTooLarge(425, 300);
    act.Should().Throw<ModelContextProtocol.McpException>()
       .Which.Message.Should().Contain("graph_too_large")
       .And.Contain("425")
       .And.Contain("300");
}

[Fact]
public void InvalidRenderOption_throws_McpException_with_option_and_message()
{
    var act = () => ToolError.InvalidRenderOption("format", "svg", "Use one of mermaid, dot.");
    act.Should().Throw<ModelContextProtocol.McpException>()
       .Which.Message.Should().Contain("invalid_render_option")
       .And.Contain("format")
       .And.Contain("svg")
       .And.Contain("mermaid");
}
```

- [ ] **Step 3: Run — they fail**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaErrorCodeTests"
```

Expected: 3 new failures (`ToolError.ProcedureNotFound` / `GraphTooLarge` / `InvalidRenderOption` don't exist).

- [ ] **Step 4: Add the codes**

Modify `src/mcpOffice/ErrorCode.cs` — add three constants under the existing list:

```csharp
public const string ProcedureNotFound = "procedure_not_found";
public const string GraphTooLarge = "graph_too_large";
public const string InvalidRenderOption = "invalid_render_option";
```

- [ ] **Step 5: Add the helpers**

Modify `src/mcpOffice/ToolError.cs` — add three methods alongside the existing helpers:

```csharp
public static Exception ProcedureNotFound(string procedureName, IEnumerable<string> available) =>
    Throw(ErrorCode.ProcedureNotFound, $"Procedure not found: {procedureName}. Available procedures: {string.Join(", ", available)}");

public static Exception GraphTooLarge(int actualNodeCount, int maxNodes) =>
    Throw(ErrorCode.GraphTooLarge, $"Filtered call graph has {actualNodeCount} nodes, which exceeds maxNodes={maxNodes}. Narrow the result with moduleName, procedureName, or a smaller depth.");

public static Exception InvalidRenderOption(string optionName, string value, string detail) =>
    Throw(ErrorCode.InvalidRenderOption, $"Invalid value for {optionName}: '{value}'. {detail}");
```

- [ ] **Step 6: Run — they pass**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaErrorCodeTests"
```

Expected: all VbaErrorCodeTests green.

- [ ] **Step 7: Commit**

```bash
git add src/mcpOffice/ErrorCode.cs src/mcpOffice/ToolError.cs tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs
git commit -m "feat: add procedure_not_found / graph_too_large / invalid_render_option codes"
```

---

# Phase 2 — DTOs

### Task 3: Add `CallgraphNode`, `CallgraphEdge`, `FilteredCallgraph` records

**Files:**
- Create: `src/mcpOffice/Models/CallgraphNode.cs`
- Create: `src/mcpOffice/Models/CallgraphEdge.cs`
- Create: `src/mcpOffice/Models/FilteredCallgraph.cs`

These DTOs are internal plumbing between the filter and the renderers; no test exercises them directly (they're exercised through the filter and renderer tests). Build cleanness is the gate.

- [ ] **Step 1: Create `CallgraphNode.cs`**

```csharp
namespace McpOffice.Models;

/// <summary>
/// A node in a filtered call graph. <see cref="Id"/> is the canonical FQN
/// (e.g., "mdlScreeningDB.ReadExports") for resolved procedures, or "__ext__&lt;name&gt;"
/// for unresolved/external callees. Renderers are responsible for mangling Id
/// into format-safe identifiers.
/// </summary>
public sealed record CallgraphNode(
    string Id,
    string Label,
    string? Module,             // null for external nodes
    bool IsEventHandler,
    bool IsOrphan,
    bool IsExternal);
```

- [ ] **Step 2: Create `CallgraphEdge.cs`**

```csharp
namespace McpOffice.Models;

public sealed record CallgraphEdge(
    string FromId,
    string ToId,
    bool Resolved);
```

- [ ] **Step 3: Create `FilteredCallgraph.cs`**

```csharp
namespace McpOffice.Models;

public sealed record FilteredCallgraph(
    IReadOnlyList<CallgraphNode> Nodes,
    IReadOnlyList<CallgraphEdge> Edges);
```

- [ ] **Step 4: Build**

```bash
dotnet build --nologo
```

Expected: 0 warnings, 0 errors.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Models/CallgraphNode.cs src/mcpOffice/Models/CallgraphEdge.cs src/mcpOffice/Models/FilteredCallgraph.cs
git commit -m "feat: add CallgraphNode/Edge/FilteredCallgraph DTOs"
```

---

# Phase 3 — Filter

The filter is the most logic-heavy piece. Each task adds one slice of behaviour with focused TDD coverage.

### Task 4: Filter — no-filter mode (whole workbook pass-through)

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs`

- [ ] **Step 1: Define filter options shape**

These options live as a record in the filter's namespace — not in `Models/` (they're plumbing, not part of the wire DTO).

Add at the top of the new `VbaCallgraphFilter.cs`:

```csharp
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

public sealed record CallgraphFilterOptions(
    string? ModuleName = null,
    string? ProcedureName = null,
    int Depth = 2,
    string Direction = "both",       // "callees" | "callers" | "both"
    int MaxNodes = 300);

public static class VbaCallgraphFilter
{
    public static FilteredCallgraph Apply(ExcelVbaAnalysis analysis, CallgraphFilterOptions options)
    {
        // To be implemented across Tasks 4–10.
        throw new NotImplementedException();
    }
}
```

- [ ] **Step 2: Write the failing test**

`tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs`:

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaCallgraphFilterTests
{
    // Helper: build a minimal ExcelVbaAnalysis with the given modules + edges.
    private static ExcelVbaAnalysis Analysis(
        IEnumerable<(string Module, string Kind, string Name, bool IsHandler)> procs,
        IEnumerable<(string From, string To, bool Resolved)> edges)
    {
        var byModule = procs
            .GroupBy(p => (p.Module, p.Kind))
            .Select(g => new ExcelVbaModuleAnalysis(
                g.Key.Module,
                g.Key.Kind,
                Parsed: true,
                Reason: null,
                Procedures: g.Select(p => new ExcelVbaProcedure(
                    Name: p.Name,
                    FullyQualifiedName: $"{p.Module}.{p.Name}",
                    Kind: "Sub",
                    Scope: null,
                    Parameters: Array.Empty<ExcelVbaParameter>(),
                    ReturnType: null,
                    LineStart: 1,
                    LineEnd: 2,
                    IsEventHandler: p.IsHandler,
                    EventTarget: null)).ToList()))
            .ToList();

        var callEdges = edges.Select(e => new ExcelVbaCallEdge(
            From: e.From,
            To: e.To,
            Resolved: e.Resolved,
            Site: new ExcelVbaSiteRef(
                Module: e.From.Split('.')[0],
                Procedure: e.From.Split('.')[1],
                Line: 1))).ToList();

        var procedureCount = byModule.Sum(m => m.Procedures.Count);
        var handlerCount = byModule.Sum(m => m.Procedures.Count(p => p.IsEventHandler));

        return new ExcelVbaAnalysis(
            HasVbaProject: true,
            Summary: new ExcelVbaAnalysisSummary(
                ModuleCount: byModule.Count,
                ParsedModuleCount: byModule.Count,
                UnparsedModuleCount: 0,
                ProcedureCount: procedureCount,
                EventHandlerCount: handlerCount,
                CallEdgeCount: callEdges.Count,
                ObjectModelReferenceCount: 0,
                DependencyCount: 0),
            Modules: byModule,
            CallGraph: callEdges,
            References: null);
    }

    [Fact]
    public void No_filter_returns_every_procedure_as_a_node()
    {
        var a = Analysis(
            procs: new[]
            {
                ("ModA", "standardModule", "P1", false),
                ("ModA", "standardModule", "P2", false),
                ("ModB", "standardModule", "Q1", false),
            },
            edges: new[]
            {
                ("ModA.P1", "ModA.P2", true),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

        Assert.Equal(3, result.Nodes.Count);
        Assert.Contains(result.Nodes, n => n.Id == "ModA.P1");
        Assert.Contains(result.Nodes, n => n.Id == "ModA.P2");
        Assert.Contains(result.Nodes, n => n.Id == "ModB.Q1");
        Assert.Single(result.Edges);
        Assert.Equal("ModA.P1", result.Edges[0].FromId);
        Assert.Equal("ModA.P2", result.Edges[0].ToId);
        Assert.True(result.Edges[0].Resolved);
    }

    [Fact]
    public void No_vba_project_returns_empty()
    {
        var empty = new ExcelVbaAnalysis(
            HasVbaProject: false,
            Summary: new ExcelVbaAnalysisSummary(0, 0, 0, 0, 0, 0, 0, 0),
            Modules: null,
            CallGraph: null,
            References: null);

        var result = VbaCallgraphFilter.Apply(empty, new CallgraphFilterOptions());

        Assert.Empty(result.Nodes);
        Assert.Empty(result.Edges);
    }
}
```

- [ ] **Step 3: Run — fails**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: `NotImplementedException`.

- [ ] **Step 4: Implement no-filter pass-through**

Replace the body of `VbaCallgraphFilter.Apply`:

```csharp
public static FilteredCallgraph Apply(ExcelVbaAnalysis analysis, CallgraphFilterOptions options)
{
    if (!analysis.HasVbaProject || analysis.Modules is null)
        return new FilteredCallgraph(Array.Empty<CallgraphNode>(), Array.Empty<CallgraphEdge>());

    var nodes = new List<CallgraphNode>();
    foreach (var m in analysis.Modules)
    {
        if (!m.Parsed) continue;
        foreach (var p in m.Procedures)
        {
            nodes.Add(new CallgraphNode(
                Id: p.FullyQualifiedName,
                Label: p.Name,
                Module: m.Name,
                IsEventHandler: p.IsEventHandler,
                IsOrphan: false,           // classified in a later task
                IsExternal: false));
        }
    }

    var edges = (analysis.CallGraph ?? Array.Empty<ExcelVbaCallEdge>())
        .Select(e => new CallgraphEdge(e.From, e.To, e.Resolved))
        .ToList();

    return new FilteredCallgraph(nodes, edges);
}
```

- [ ] **Step 5: Run — passes**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: both tests green.

- [ ] **Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs
git commit -m "feat: VbaCallgraphFilter — no-filter pass-through"
```

---

### Task 5: Filter — `moduleName` direct-neighbour expansion + `module_not_found`

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs`

- [ ] **Step 1: Write failing tests**

Append to `VbaCallgraphFilterTests.cs`:

```csharp
[Fact]
public void Module_filter_includes_module_procedures_and_direct_neighbours()
{
    // ModA calls ModB, ModC stands alone.
    var a = Analysis(
        procs: new[]
        {
            ("ModA", "standardModule", "P1", false),
            ("ModB", "standardModule", "Q1", false),
            ("ModC", "standardModule", "R1", false),
        },
        edges: new[]
        {
            ("ModA.P1", "ModB.Q1", true),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "ModA"));

    // Expect ModA.P1 (in module) + ModB.Q1 (direct neighbour). ModC.R1 dropped.
    Assert.Equal(2, result.Nodes.Count);
    Assert.Contains(result.Nodes, n => n.Id == "ModA.P1");
    Assert.Contains(result.Nodes, n => n.Id == "ModB.Q1");
    Assert.DoesNotContain(result.Nodes, n => n.Id == "ModC.R1");
    Assert.Single(result.Edges);
}

[Fact]
public void Module_filter_pulls_in_callers_too()
{
    // ModB.Q1 calls ModA.P1 (caller direction).
    var a = Analysis(
        procs: new[]
        {
            ("ModA", "standardModule", "P1", false),
            ("ModB", "standardModule", "Q1", false),
        },
        edges: new[]
        {
            ("ModB.Q1", "ModA.P1", true),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "ModA"));

    Assert.Equal(2, result.Nodes.Count);
    Assert.Single(result.Edges);
}

[Fact]
public void Module_filter_is_case_insensitive()
{
    var a = Analysis(
        procs: new[] { ("ModA", "standardModule", "P1", false) },
        edges: Array.Empty<(string, string, bool)>());

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "moda"));

    Assert.Single(result.Nodes);
}

[Fact]
public void Module_filter_unknown_throws_module_not_found()
{
    var a = Analysis(
        procs: new[] { ("ModA", "standardModule", "P1", false) },
        edges: Array.Empty<(string, string, bool)>());

    var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "Nope"));
    var ex = Assert.Throws<ModelContextProtocol.McpException>(act);
    Assert.Contains("module_not_found", ex.Message);
    Assert.Contains("ModA", ex.Message);
}
```

- [ ] **Step 2: Run — they fail**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: 4 new failures.

- [ ] **Step 3: Implement module filter**

Modify `VbaCallgraphFilter.Apply` — after the early-out for `HasVbaProject = false`, resolve and apply the module filter:

```csharp
// Resolve module filter (case-insensitive) — produces canonical casing for downstream comparisons.
string? moduleFilter = null;
if (!string.IsNullOrWhiteSpace(options.ModuleName))
{
    var match = analysis.Modules.FirstOrDefault(m =>
        string.Equals(m.Name, options.ModuleName, StringComparison.OrdinalIgnoreCase));
    if (match is null)
        throw ToolError.ModuleNotFound(options.ModuleName, analysis.Modules.Select(m => m.Name));
    moduleFilter = match.Name;
}

// Build the full procedure-node set first (every parsed procedure across all modules).
var allNodesById = new Dictionary<string, CallgraphNode>();
foreach (var m in analysis.Modules)
{
    if (!m.Parsed) continue;
    foreach (var p in m.Procedures)
    {
        allNodesById[p.FullyQualifiedName] = new CallgraphNode(
            Id: p.FullyQualifiedName,
            Label: p.Name,
            Module: m.Name,
            IsEventHandler: p.IsEventHandler,
            IsOrphan: false,
            IsExternal: false);
    }
}

var allEdges = analysis.CallGraph ?? Array.Empty<ExcelVbaCallEdge>();

if (moduleFilter is null)
{
    // No-filter mode: return all procedure nodes + all edges between known nodes.
    var passThruEdges = allEdges
        .Where(e => allNodesById.ContainsKey(e.From) && allNodesById.ContainsKey(e.To))
        .Select(e => new CallgraphEdge(e.From, e.To, e.Resolved))
        .ToList();
    return new FilteredCallgraph(allNodesById.Values.ToList(), passThruEdges);
}

// Module-only mode: seed = procs in module; expand one hop both directions.
var moduleProcIds = allNodesById.Values
    .Where(n => n.Module == moduleFilter)
    .Select(n => n.Id)
    .ToHashSet();

var survivingIds = new HashSet<string>(moduleProcIds);
foreach (var e in allEdges)
{
    if (moduleProcIds.Contains(e.From) && allNodesById.ContainsKey(e.To))
        survivingIds.Add(e.To);
    if (moduleProcIds.Contains(e.To) && allNodesById.ContainsKey(e.From))
        survivingIds.Add(e.From);
}

var moduleNodes = survivingIds.Select(id => allNodesById[id]).ToList();
var moduleEdges = allEdges
    .Where(e => survivingIds.Contains(e.From) && survivingIds.Contains(e.To))
    .Select(e => new CallgraphEdge(e.From, e.To, e.Resolved))
    .ToList();

return new FilteredCallgraph(moduleNodes, moduleEdges);
```

Replace the previous body (the no-filter early return is now folded into the conditional above).

- [ ] **Step 4: Run — passes**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: all 6 tests green (the 2 from Task 4 plus the 4 added here).

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs
git commit -m "feat: VbaCallgraphFilter — moduleName direct-neighbour expansion"
```

---

### Task 6: Filter — focal procedure BFS (`procedureName` + `depth` + `direction`) + `procedure_not_found`

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs`

- [ ] **Step 1: Write failing tests**

Append to `VbaCallgraphFilterTests.cs`:

```csharp
[Fact]
public void Focal_procedure_callees_only_depth_1()
{
    // P1 → P2 → P3, plus an unrelated P4.
    var a = Analysis(
        procs: new[]
        {
            ("M", "standardModule", "P1", false),
            ("M", "standardModule", "P2", false),
            ("M", "standardModule", "P3", false),
            ("M", "standardModule", "P4", false),
        },
        edges: new[]
        {
            ("M.P1", "M.P2", true),
            ("M.P2", "M.P3", true),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ModuleName: "M",
        ProcedureName: "P1",
        Depth: 1,
        Direction: "callees"));

    // From P1, callees-only depth 1 → {P1, P2}. P3 needs depth 2.
    Assert.Equal(2, result.Nodes.Count);
    Assert.Contains(result.Nodes, n => n.Id == "M.P1");
    Assert.Contains(result.Nodes, n => n.Id == "M.P2");
    Assert.DoesNotContain(result.Nodes, n => n.Id == "M.P3");
    Assert.DoesNotContain(result.Nodes, n => n.Id == "M.P4");
}

[Fact]
public void Focal_procedure_callees_depth_2_pulls_in_grandchildren()
{
    var a = Analysis(
        procs: new[]
        {
            ("M", "standardModule", "P1", false),
            ("M", "standardModule", "P2", false),
            ("M", "standardModule", "P3", false),
        },
        edges: new[]
        {
            ("M.P1", "M.P2", true),
            ("M.P2", "M.P3", true),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ModuleName: "M",
        ProcedureName: "P1",
        Depth: 2,
        Direction: "callees"));

    Assert.Equal(3, result.Nodes.Count);
}

[Fact]
public void Focal_procedure_callers_walks_inbound_edges()
{
    // P1 ← P2 ← P3, P4 unrelated.
    var a = Analysis(
        procs: new[]
        {
            ("M", "standardModule", "P1", false),
            ("M", "standardModule", "P2", false),
            ("M", "standardModule", "P3", false),
            ("M", "standardModule", "P4", false),
        },
        edges: new[]
        {
            ("M.P2", "M.P1", true),
            ("M.P3", "M.P2", true),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ModuleName: "M",
        ProcedureName: "P1",
        Depth: 2,
        Direction: "callers"));

    Assert.Equal(3, result.Nodes.Count);  // P1, P2, P3
    Assert.DoesNotContain(result.Nodes, n => n.Id == "M.P4");
}

[Fact]
public void Focal_procedure_both_unions_callees_and_callers()
{
    // P0 → P1 → P2, P3 → P1.
    var a = Analysis(
        procs: new[]
        {
            ("M", "standardModule", "P0", false),
            ("M", "standardModule", "P1", false),
            ("M", "standardModule", "P2", false),
            ("M", "standardModule", "P3", false),
        },
        edges: new[]
        {
            ("M.P0", "M.P1", true),
            ("M.P1", "M.P2", true),
            ("M.P3", "M.P1", true),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ModuleName: "M",
        ProcedureName: "P1",
        Depth: 1,
        Direction: "both"));

    // {P1} + callees(P1)={P2} + callers(P1)={P0,P3}
    Assert.Equal(4, result.Nodes.Count);
}

[Fact]
public void Focal_procedure_depth_zero_returns_just_the_seed()
{
    var a = Analysis(
        procs: new[]
        {
            ("M", "standardModule", "P1", false),
            ("M", "standardModule", "P2", false),
        },
        edges: new[] { ("M.P1", "M.P2", true) });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ModuleName: "M",
        ProcedureName: "P1",
        Depth: 0,
        Direction: "both"));

    Assert.Single(result.Nodes);
    Assert.Equal("M.P1", result.Nodes[0].Id);
    Assert.Empty(result.Edges);
}

[Fact]
public void ProcedureName_unknown_throws_procedure_not_found()
{
    var a = Analysis(
        procs: new[]
        {
            ("M", "standardModule", "P1", false),
            ("M", "standardModule", "P2", false),
        },
        edges: Array.Empty<(string, string, bool)>());

    var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ModuleName: "M",
        ProcedureName: "Nope"));
    var ex = Assert.Throws<ModelContextProtocol.McpException>(act);
    Assert.Contains("procedure_not_found", ex.Message);
    Assert.Contains("Nope", ex.Message);
    Assert.Contains("P1", ex.Message);
    Assert.Contains("P2", ex.Message);
}

[Fact]
public void ProcedureName_is_case_insensitive_within_module()
{
    var a = Analysis(
        procs: new[] { ("M", "standardModule", "ReadExports", false) },
        edges: Array.Empty<(string, string, bool)>());

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ModuleName: "M",
        ProcedureName: "readexports",
        Depth: 0));

    Assert.Single(result.Nodes);
}

[Fact]
public void Direction_unknown_value_throws_invalid_render_option()
{
    var a = Analysis(
        procs: new[] { ("M", "standardModule", "P1", false) },
        edges: Array.Empty<(string, string, bool)>());

    var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ModuleName: "M",
        ProcedureName: "P1",
        Direction: "sideways"));
    var ex = Assert.Throws<ModelContextProtocol.McpException>(act);
    Assert.Contains("invalid_render_option", ex.Message);
    Assert.Contains("direction", ex.Message);
    Assert.Contains("sideways", ex.Message);
}
```

- [ ] **Step 2: Run — they fail**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: 8 new failures.

- [ ] **Step 3: Replace the implementation with focal-aware version**

Replace the body of `VbaCallgraphFilter.Apply` so that when `ProcedureName` is set, BFS-from-focal runs *after* the module is resolved:

```csharp
public static FilteredCallgraph Apply(ExcelVbaAnalysis analysis, CallgraphFilterOptions options)
{
    if (!analysis.HasVbaProject || analysis.Modules is null)
        return new FilteredCallgraph(Array.Empty<CallgraphNode>(), Array.Empty<CallgraphEdge>());

    string? moduleFilter = null;
    if (!string.IsNullOrWhiteSpace(options.ModuleName))
    {
        var match = analysis.Modules.FirstOrDefault(m =>
            string.Equals(m.Name, options.ModuleName, StringComparison.OrdinalIgnoreCase));
        if (match is null)
            throw ToolError.ModuleNotFound(options.ModuleName, analysis.Modules.Select(m => m.Name));
        moduleFilter = match.Name;
    }

    // Build the full procedure-node set across all parsed modules.
    var allNodesById = new Dictionary<string, CallgraphNode>();
    foreach (var m in analysis.Modules)
    {
        if (!m.Parsed) continue;
        foreach (var p in m.Procedures)
        {
            allNodesById[p.FullyQualifiedName] = new CallgraphNode(
                Id: p.FullyQualifiedName,
                Label: p.Name,
                Module: m.Name,
                IsEventHandler: p.IsEventHandler,
                IsOrphan: false,
                IsExternal: false);
        }
    }

    var allEdges = analysis.CallGraph ?? Array.Empty<ExcelVbaCallEdge>();

    // Branch 1: focal-procedure BFS.
    if (!string.IsNullOrWhiteSpace(options.ProcedureName))
    {
        if (moduleFilter is null)
            throw ToolError.InvalidRenderOption(
                "procedureName", options.ProcedureName,
                "procedureName requires moduleName — bare procedure names aren't unique.");

        var moduleProcs = analysis.Modules
            .Single(m => m.Name == moduleFilter)
            .Procedures;
        var focalMatch = moduleProcs.FirstOrDefault(p =>
            string.Equals(p.Name, options.ProcedureName, StringComparison.OrdinalIgnoreCase));
        if (focalMatch is null)
            throw ToolError.ProcedureNotFound(options.ProcedureName, moduleProcs.Select(p => p.Name));

        var focalId = focalMatch.FullyQualifiedName;
        var followCallees = options.Direction is "callees" or "both";
        var followCallers = options.Direction is "callers" or "both";
        if (!followCallees && !followCallers)
            throw ToolError.InvalidRenderOption(
                "direction", options.Direction,
                "Use one of callees, callers, both.");

        var visited = new HashSet<string> { focalId };
        var frontier = new HashSet<string> { focalId };
        for (var hop = 0; hop < options.Depth; hop++)
        {
            var next = new HashSet<string>();
            foreach (var e in allEdges)
            {
                if (followCallees && frontier.Contains(e.From) && !visited.Contains(e.To)
                    && (allNodesById.ContainsKey(e.To) || !e.Resolved))
                    next.Add(e.To);
                if (followCallers && frontier.Contains(e.To) && !visited.Contains(e.From)
                    && allNodesById.ContainsKey(e.From))
                    next.Add(e.From);
            }
            if (next.Count == 0) break;
            foreach (var id in next) visited.Add(id);
            frontier = next;
        }

        var bfsNodes = visited
            .Where(allNodesById.ContainsKey)
            .Select(id => allNodesById[id])
            .ToList();
        var bfsEdges = allEdges
            .Where(e => visited.Contains(e.From) && visited.Contains(e.To))
            .Select(e => new CallgraphEdge(e.From, e.To, e.Resolved))
            .ToList();
        return new FilteredCallgraph(bfsNodes, bfsEdges);
    }

    // Branch 2: moduleName-only direct-neighbour expansion.
    if (moduleFilter is not null)
    {
        var moduleProcIds = allNodesById.Values
            .Where(n => n.Module == moduleFilter)
            .Select(n => n.Id)
            .ToHashSet();
        var survivingIds = new HashSet<string>(moduleProcIds);
        foreach (var e in allEdges)
        {
            if (moduleProcIds.Contains(e.From) && allNodesById.ContainsKey(e.To))
                survivingIds.Add(e.To);
            if (moduleProcIds.Contains(e.To) && allNodesById.ContainsKey(e.From))
                survivingIds.Add(e.From);
        }

        var moduleNodes = survivingIds.Select(id => allNodesById[id]).ToList();
        var moduleEdges = allEdges
            .Where(e => survivingIds.Contains(e.From) && survivingIds.Contains(e.To))
            .Select(e => new CallgraphEdge(e.From, e.To, e.Resolved))
            .ToList();
        return new FilteredCallgraph(moduleNodes, moduleEdges);
    }

    // Branch 3: no filter — return everything.
    var passThruEdges = allEdges
        .Where(e => allNodesById.ContainsKey(e.From) && allNodesById.ContainsKey(e.To))
        .Select(e => new CallgraphEdge(e.From, e.To, e.Resolved))
        .ToList();
    return new FilteredCallgraph(allNodesById.Values.ToList(), passThruEdges);
}
```

- [ ] **Step 4: Run — they pass**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: all green.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs
git commit -m "feat: VbaCallgraphFilter — focal-procedure BFS with depth and direction"
```

---

### Task 7: Filter — `procedureName` without `moduleName` throws `invalid_render_option`

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs`

The implementation already handles this (Task 6 added the `InvalidRenderOption` throw inside the focal branch when `moduleFilter is null`). We only need the test that pins it.

- [ ] **Step 1: Write the test**

Append to `VbaCallgraphFilterTests.cs`:

```csharp
[Fact]
public void ProcedureName_without_moduleName_throws_invalid_render_option()
{
    var a = Analysis(
        procs: new[] { ("M", "standardModule", "P1", false) },
        edges: Array.Empty<(string, string, bool)>());

    var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
        ProcedureName: "P1"));
    var ex = Assert.Throws<ModelContextProtocol.McpException>(act);
    Assert.Contains("invalid_render_option", ex.Message);
    Assert.Contains("procedureName", ex.Message);
    Assert.Contains("requires moduleName", ex.Message);
}
```

- [ ] **Step 2: Run — passes**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: green (implementation from Task 6 already covers this).

- [ ] **Step 3: Commit**

```bash
git add tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs
git commit -m "test: VbaCallgraphFilter — procedureName-without-moduleName guard"
```

---

### Task 8: Filter — external (unresolved) callees become deduplicated `<external>` nodes

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs`

The current filter drops unresolved-target edges silently because their `To` is missing from `allNodesById`. Fix: synthesise one `<external>` node per distinct unresolved callee, deduplicated.

- [ ] **Step 1: Write failing tests**

Append to `VbaCallgraphFilterTests.cs`:

```csharp
[Fact]
public void Unresolved_callees_become_single_external_node_per_name()
{
    // Two procedures both call MsgBox (unresolved) — one external node, two edges into it.
    var a = Analysis(
        procs: new[]
        {
            ("M", "standardModule", "P1", false),
            ("M", "standardModule", "P2", false),
        },
        edges: new[]
        {
            ("M.P1", "MsgBox", false),
            ("M.P2", "MsgBox", false),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

    var externals = result.Nodes.Where(n => n.IsExternal).ToList();
    Assert.Single(externals);
    Assert.Equal("MsgBox", externals[0].Label);
    Assert.True(externals[0].IsExternal);
    Assert.Null(externals[0].Module);

    // Both inbound edges land on the same external node id.
    var externalEdges = result.Edges.Where(e => e.ToId == externals[0].Id).ToList();
    Assert.Equal(2, externalEdges.Count);
    Assert.All(externalEdges, e => Assert.False(e.Resolved));
}

[Fact]
public void Distinct_unresolved_names_get_distinct_external_nodes()
{
    var a = Analysis(
        procs: new[] { ("M", "standardModule", "P1", false) },
        edges: new[]
        {
            ("M.P1", "MsgBox", false),
            ("M.P1", "CreateObject", false),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

    var externals = result.Nodes.Where(n => n.IsExternal).ToList();
    Assert.Equal(2, externals.Count);
    Assert.Contains(externals, e => e.Label == "MsgBox");
    Assert.Contains(externals, e => e.Label == "CreateObject");
}

[Fact]
public void External_nodes_appear_only_when_caller_is_in_filtered_set()
{
    // ModA.P1 calls MsgBox; ModB.Q1 also calls MsgBox. Filter to ModA.
    var a = Analysis(
        procs: new[]
        {
            ("ModA", "standardModule", "P1", false),
            ("ModB", "standardModule", "Q1", false),
        },
        edges: new[]
        {
            ("ModA.P1", "MsgBox", false),
            ("ModB.Q1", "MsgBox", false),
        });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "ModA"));

    // Only one external (Q1 is filtered out, so its MsgBox edge can't pull the external in).
    var externals = result.Nodes.Where(n => n.IsExternal).ToList();
    Assert.Single(externals);
    var externalEdges = result.Edges.Where(e => e.ToId == externals[0].Id).ToList();
    Assert.Single(externalEdges);
    Assert.Equal("ModA.P1", externalEdges[0].FromId);
}
```

- [ ] **Step 2: Run — they fail**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: 3 new failures (externals currently dropped).

- [ ] **Step 3: Add a helper for external node IDs**

Inside `VbaCallgraphFilter`, add a private helper at the bottom of the class:

```csharp
private const string ExternalIdPrefix = "__ext__";

private static string ExternalId(string calleeName) => ExternalIdPrefix + calleeName;
```

- [ ] **Step 4: Update each branch to emit external nodes**

Replace the three branches' edge-handling so that unresolved edges synthesise external nodes scoped to surviving callers. The cleanest factoring is a small helper:

Add inside `VbaCallgraphFilter`:

```csharp
private static (List<CallgraphNode> Nodes, List<CallgraphEdge> Edges) BuildOutput(
    HashSet<string> survivingProcIds,
    Dictionary<string, CallgraphNode> allNodesById,
    IReadOnlyList<ExcelVbaCallEdge> allEdges)
{
    var outNodes = survivingProcIds
        .Where(allNodesById.ContainsKey)
        .Select(id => allNodesById[id])
        .ToList();

    var externalIds = new Dictionary<string, CallgraphNode>(StringComparer.Ordinal);
    var outEdges = new List<CallgraphEdge>();

    foreach (var e in allEdges)
    {
        var fromIsProc = survivingProcIds.Contains(e.From);
        var toIsProc = allNodesById.ContainsKey(e.To) && survivingProcIds.Contains(e.To);

        if (fromIsProc && toIsProc)
        {
            outEdges.Add(new CallgraphEdge(e.From, e.To, e.Resolved));
        }
        else if (fromIsProc && !e.Resolved)
        {
            // Unresolved external — synthesise / reuse external node.
            var extId = ExternalId(e.To);
            if (!externalIds.ContainsKey(extId))
            {
                externalIds[extId] = new CallgraphNode(
                    Id: extId,
                    Label: e.To,
                    Module: null,
                    IsEventHandler: false,
                    IsOrphan: false,
                    IsExternal: true);
            }
            outEdges.Add(new CallgraphEdge(e.From, extId, false));
        }
        // else: edge dropped (To unknown but Resolved=true, or From not surviving).
    }

    outNodes.AddRange(externalIds.Values);
    return (outNodes, outEdges);
}
```

Now replace the three "build the output" tails of each branch with calls to this helper.

**Branch 1 (focal):** after the BFS expansion, replace the final two `bfs*` lines with:

```csharp
var (bfsNodes, bfsEdges) = BuildOutput(
    survivingProcIds: visited.Where(allNodesById.ContainsKey).ToHashSet(),
    allNodesById,
    allEdges);
return new FilteredCallgraph(bfsNodes, bfsEdges);
```

**Branch 2 (module-only):** replace the final `moduleNodes` / `moduleEdges` lines with:

```csharp
var (moduleNodes, moduleEdges) = BuildOutput(survivingIds, allNodesById, allEdges);
return new FilteredCallgraph(moduleNodes, moduleEdges);
```

**Branch 3 (no filter):** replace `passThruEdges` and the return with:

```csharp
var allProcIds = allNodesById.Keys.ToHashSet();
var (allNodes, allEdgesOut) = BuildOutput(allProcIds, allNodesById, allEdges);
return new FilteredCallgraph(allNodes, allEdgesOut);
```

- [ ] **Step 5: Update Task 4's earlier test**

`No_filter_returns_every_procedure_as_a_node` was written before externals. The test as written has only resolved edges, so its assertions still hold — leave it alone.

`Module_filter_includes_module_procedures_and_direct_neighbours` likewise — only resolved edges. Leave alone.

- [ ] **Step 6: Run — all green**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: every test green.

- [ ] **Step 7: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs
git commit -m "feat: VbaCallgraphFilter — external nodes deduplicated per callee name"
```

---

### Task 9: Filter — orphan classification

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs`

A node is an orphan **after filtering** when it has zero in/out edges among surviving edges *and* `IsEventHandler=false`. Re-stamp `IsOrphan` on each surviving procedure node before returning.

- [ ] **Step 1: Write failing tests**

```csharp
[Fact]
public void Orphan_procedure_with_no_edges_marked_isOrphan()
{
    var a = Analysis(
        procs: new[]
        {
            ("M", "standardModule", "Connected1", false),
            ("M", "standardModule", "Connected2", false),
            ("M", "standardModule", "Lonely", false),
        },
        edges: new[] { ("M.Connected1", "M.Connected2", true) });

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

    var lonely = Assert.Single(result.Nodes.Where(n => n.Id == "M.Lonely"));
    Assert.True(lonely.IsOrphan);
    Assert.False(result.Nodes.Single(n => n.Id == "M.Connected1").IsOrphan);
}

[Fact]
public void Event_handler_with_no_edges_is_NOT_marked_orphan()
{
    var a = Analysis(
        procs: new[]
        {
            ("ThisWorkbook", "documentModule", "Workbook_Open", true),
        },
        edges: Array.Empty<(string, string, bool)>());

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

    var node = Assert.Single(result.Nodes);
    Assert.True(node.IsEventHandler);
    Assert.False(node.IsOrphan);
}

[Fact]
public void Orphan_classification_is_per_filtered_view_not_per_workbook()
{
    // ModA.P1 calls ModB.Q1. Filter to ModA: Q1 is pulled in but, in the filtered view,
    // P1 has out-edge to Q1 so P1 isn't orphan. Q1 has in-edge from P1 so Q1 isn't orphan.
    // Now consider an extra ModA.X with no edges — it should be orphan in both views.
    var a = Analysis(
        procs: new[]
        {
            ("ModA", "standardModule", "P1", false),
            ("ModA", "standardModule", "X", false),
            ("ModB", "standardModule", "Q1", false),
        },
        edges: new[] { ("ModA.P1", "ModB.Q1", true) });

    var moduleResult = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "ModA"));
    var noFilterResult = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

    Assert.True(moduleResult.Nodes.Single(n => n.Id == "ModA.X").IsOrphan);
    Assert.True(noFilterResult.Nodes.Single(n => n.Id == "ModA.X").IsOrphan);
    Assert.False(noFilterResult.Nodes.Single(n => n.Id == "ModA.P1").IsOrphan);
}
```

- [ ] **Step 2: Run — they fail**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

Expected: failures (orphan classification not yet implemented).

- [ ] **Step 3: Implement**

Modify `BuildOutput` so that after edges are computed, each procedure node is re-stamped with the correct `IsOrphan` flag based on the surviving edge set. Replace `BuildOutput` with:

```csharp
private static (List<CallgraphNode> Nodes, List<CallgraphEdge> Edges) BuildOutput(
    HashSet<string> survivingProcIds,
    Dictionary<string, CallgraphNode> allNodesById,
    IReadOnlyList<ExcelVbaCallEdge> allEdges)
{
    var externalIds = new Dictionary<string, CallgraphNode>(StringComparer.Ordinal);
    var outEdges = new List<CallgraphEdge>();

    foreach (var e in allEdges)
    {
        var fromIsProc = survivingProcIds.Contains(e.From);
        var toIsProc = allNodesById.ContainsKey(e.To) && survivingProcIds.Contains(e.To);

        if (fromIsProc && toIsProc)
        {
            outEdges.Add(new CallgraphEdge(e.From, e.To, e.Resolved));
        }
        else if (fromIsProc && !e.Resolved)
        {
            var extId = ExternalId(e.To);
            if (!externalIds.ContainsKey(extId))
            {
                externalIds[extId] = new CallgraphNode(
                    Id: extId,
                    Label: e.To,
                    Module: null,
                    IsEventHandler: false,
                    IsOrphan: false,
                    IsExternal: true);
            }
            outEdges.Add(new CallgraphEdge(e.From, extId, false));
        }
    }

    // Compute degree per surviving node id.
    var degree = new Dictionary<string, int>();
    foreach (var e in outEdges)
    {
        degree[e.FromId] = degree.GetValueOrDefault(e.FromId) + 1;
        degree[e.ToId] = degree.GetValueOrDefault(e.ToId) + 1;
    }

    var outNodes = new List<CallgraphNode>(survivingProcIds.Count + externalIds.Count);
    foreach (var id in survivingProcIds)
    {
        if (!allNodesById.TryGetValue(id, out var node)) continue;
        var isOrphan = !node.IsEventHandler && !degree.ContainsKey(id);
        outNodes.Add(node with { IsOrphan = isOrphan });
    }
    outNodes.AddRange(externalIds.Values);

    return (outNodes, outEdges);
}
```

- [ ] **Step 4: Run — all green**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs
git commit -m "feat: VbaCallgraphFilter — orphan classification per filtered view"
```

---

### Task 10: Filter — `maxNodes` cap throws `graph_too_large`

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs`

- [ ] **Step 1: Write failing tests**

```csharp
[Fact]
public void Exceeds_maxNodes_throws_graph_too_large()
{
    var procs = Enumerable.Range(0, 5)
        .Select(i => ("M", "standardModule", $"P{i}", false))
        .ToArray();
    var a = Analysis(procs, Array.Empty<(string, string, bool)>());

    var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(MaxNodes: 3));
    var ex = Assert.Throws<ModelContextProtocol.McpException>(act);
    Assert.Contains("graph_too_large", ex.Message);
    Assert.Contains("5", ex.Message);
    Assert.Contains("3", ex.Message);
}

[Fact]
public void Equal_to_maxNodes_does_not_throw()
{
    var procs = Enumerable.Range(0, 3)
        .Select(i => ("M", "standardModule", $"P{i}", false))
        .ToArray();
    var a = Analysis(procs, Array.Empty<(string, string, bool)>());

    var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(MaxNodes: 3));
    Assert.Equal(3, result.Nodes.Count);
}
```

- [ ] **Step 2: Run — they fail**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

- [ ] **Step 3: Implement**

In `VbaCallgraphFilter.Apply`, just before each `return new FilteredCallgraph(...)`, check the cap. The cleanest approach is to centralise the check in a tiny tail-call helper. Add at the bottom of the class:

```csharp
private static FilteredCallgraph Cap(FilteredCallgraph graph, int maxNodes)
{
    if (graph.Nodes.Count > maxNodes)
        throw ToolError.GraphTooLarge(graph.Nodes.Count, maxNodes);
    return graph;
}
```

Wrap each return — there are three returns (one per branch) plus the empty-vba early return. Empty case is fine to skip; wrap the three real returns:

```csharp
return Cap(new FilteredCallgraph(bfsNodes, bfsEdges), options.MaxNodes);
// ...
return Cap(new FilteredCallgraph(moduleNodes, moduleEdges), options.MaxNodes);
// ...
return Cap(new FilteredCallgraph(allNodes, allEdgesOut), options.MaxNodes);
```

- [ ] **Step 4: Run — all green**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~VbaCallgraphFilterTests"
```

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaCallgraphFilter.cs tests/mcpOffice.Tests/Excel/Vba/VbaCallgraphFilterTests.cs
git commit -m "feat: VbaCallgraphFilter — maxNodes cap with graph_too_large"
```

---

# Phase 4 — Renderers

### Task 11: `ICallgraphRenderer` interface + `MermaidCallgraphRenderer` (basics)

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/Rendering/ICallgraphRenderer.cs`
- Create: `src/mcpOffice/Services/Excel/Vba/Rendering/MermaidCallgraphRenderer.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/Rendering/MermaidCallgraphRendererTests.cs`

- [ ] **Step 1: Define render options + interface**

`src/mcpOffice/Services/Excel/Vba/Rendering/ICallgraphRenderer.cs`:

```csharp
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba.Rendering;

public sealed record CallgraphRenderOptions(
    string Layout = "clustered");   // "clustered" | "flat"

public interface ICallgraphRenderer
{
    string Render(FilteredCallgraph graph, CallgraphRenderOptions options);
}
```

- [ ] **Step 2: Write failing test — empty graph**

`tests/mcpOffice.Tests/Excel/Vba/Rendering/MermaidCallgraphRendererTests.cs`:

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba.Rendering;

namespace McpOffice.Tests.Excel.Vba.Rendering;

public class MermaidCallgraphRendererTests
{
    private static readonly MermaidCallgraphRenderer R = new();

    [Fact]
    public void Empty_graph_emits_flowchart_header()
    {
        var output = R.Render(
            new FilteredCallgraph(Array.Empty<CallgraphNode>(), Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions());

        Assert.StartsWith("flowchart TD", output);
    }

    [Fact]
    public void Single_node_clustered_wraps_in_subgraph()
    {
        var node = new CallgraphNode("M.P1", "P1", "M", IsEventHandler: false, IsOrphan: true, IsExternal: false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "clustered"));

        Assert.Contains("subgraph M", output);
        Assert.Contains("end", output);
        // Mangled id (no dots), label preserved.
        Assert.Contains("M_P1", output);
        Assert.Contains("[P1]", output);
    }

    [Fact]
    public void Single_node_flat_no_subgraphs()
    {
        var node = new CallgraphNode("M.P1", "P1", "M", false, true, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.DoesNotContain("subgraph", output);
        Assert.Contains("M_P1", output);
        Assert.Contains("[M.P1]", output);   // FQN as label in flat mode
    }

    [Fact]
    public void Edge_resolved_emits_solid_arrow()
    {
        var p1 = new CallgraphNode("M.P1", "P1", "M", false, false, false);
        var p2 = new CallgraphNode("M.P2", "P2", "M", false, false, false);
        var edge = new CallgraphEdge("M.P1", "M.P2", Resolved: true);

        var output = R.Render(
            new FilteredCallgraph(new[] { p1, p2 }, new[] { edge }),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Matches(@"M_P1\s*-->\s*M_P2", output);
    }

    [Fact]
    public void Edge_unresolved_emits_dashed_arrow()
    {
        var p1 = new CallgraphNode("M.P1", "P1", "M", false, false, false);
        var ext = new CallgraphNode("__ext__MsgBox", "MsgBox", null, false, false, true);
        var edge = new CallgraphEdge("M.P1", "__ext__MsgBox", Resolved: false);

        var output = R.Render(
            new FilteredCallgraph(new[] { p1, ext }, new[] { edge }),
            new CallgraphRenderOptions(Layout: "flat"));

        // Mermaid dashed-link syntax: -..->
        Assert.Matches(@"M_P1\s*-\.->\s*__ext__MsgBox", output);
    }
}
```

- [ ] **Step 3: Run — they fail**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~MermaidCallgraphRendererTests"
```

Expected: 5 failures (renderer doesn't exist).

- [ ] **Step 4: Implement the renderer**

`src/mcpOffice/Services/Excel/Vba/Rendering/MermaidCallgraphRenderer.cs`:

```csharp
using System.Text;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba.Rendering;

public sealed class MermaidCallgraphRenderer : ICallgraphRenderer
{
    public string Render(FilteredCallgraph graph, CallgraphRenderOptions options)
    {
        var sb = new StringBuilder();
        sb.AppendLine("flowchart TD");

        if (options.Layout == "clustered")
            EmitClustered(sb, graph);
        else
            EmitFlat(sb, graph);

        EmitEdges(sb, graph);
        EmitClassDefs(sb);

        return sb.ToString();
    }

    private static void EmitClustered(StringBuilder sb, FilteredCallgraph graph)
    {
        var grouped = graph.Nodes
            .Where(n => !n.IsExternal)
            .GroupBy(n => n.Module!)
            .OrderBy(g => g.Key, StringComparer.Ordinal);

        foreach (var group in grouped)
        {
            sb.Append("  subgraph ").AppendLine(MangleId(group.Key));
            foreach (var node in group)
            {
                sb.Append("    ");
                EmitNode(sb, node, useFqnLabel: false);
                sb.AppendLine();
            }
            sb.AppendLine("  end");
        }

        // External nodes live outside any subgraph.
        foreach (var ext in graph.Nodes.Where(n => n.IsExternal))
        {
            sb.Append("  ");
            EmitNode(sb, ext, useFqnLabel: false);
            sb.AppendLine();
        }
    }

    private static void EmitFlat(StringBuilder sb, FilteredCallgraph graph)
    {
        foreach (var node in graph.Nodes)
        {
            sb.Append("  ");
            EmitNode(sb, node, useFqnLabel: !node.IsExternal);
            sb.AppendLine();
        }
    }

    private static void EmitNode(StringBuilder sb, CallgraphNode node, bool useFqnLabel)
    {
        var id = MangleId(node.Id);
        var label = EscapeLabel(useFqnLabel ? node.Id : node.Label);

        // Shape: rounded for handlers, default rectangle otherwise.
        if (node.IsEventHandler)
            sb.Append(id).Append("([").Append(label).Append("])");
        else
            sb.Append(id).Append('[').Append(label).Append(']');

        // Class assignments.
        if (node.IsExternal) sb.Append(":::external");
        else if (node.IsEventHandler) sb.Append(":::handler");
        else if (node.IsOrphan) sb.Append(":::orphan");
    }

    private static void EmitEdges(StringBuilder sb, FilteredCallgraph graph)
    {
        foreach (var e in graph.Edges)
        {
            sb.Append("  ").Append(MangleId(e.FromId));
            sb.Append(e.Resolved ? " --> " : " -.-> ");
            sb.AppendLine(MangleId(e.ToId));
        }
    }

    private static void EmitClassDefs(StringBuilder sb)
    {
        sb.AppendLine("  classDef handler fill:#e1f5ff,stroke:#0277bd");
        sb.AppendLine("  classDef orphan stroke-dasharray:5 5");
        sb.AppendLine("  classDef external fill:#f5f5f5,stroke-dasharray:3 3");
    }

    private static string MangleId(string id)
    {
        // Mermaid IDs: alphanumeric and underscore. Replace anything else with underscore.
        var chars = id.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray();
        return new string(chars);
    }

    private static string EscapeLabel(string label)
    {
        // Quote characters that confuse Mermaid label parsing inside [].
        // Strategy: replace problematic chars with their HTML entity.
        return label
            .Replace("\"", "&quot;")
            .Replace("[", "&#91;")
            .Replace("]", "&#93;")
            .Replace("(", "&#40;")
            .Replace(")", "&#41;");
    }
}
```

- [ ] **Step 5: Run — all green**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~MermaidCallgraphRendererTests"
```

- [ ] **Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/Rendering/ICallgraphRenderer.cs src/mcpOffice/Services/Excel/Vba/Rendering/MermaidCallgraphRenderer.cs tests/mcpOffice.Tests/Excel/Vba/Rendering/MermaidCallgraphRendererTests.cs
git commit -m "feat: ICallgraphRenderer + Mermaid renderer (clusters, flat, edges, classes)"
```

---

### Task 12: Mermaid renderer — escaping reserved characters in IDs and labels

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/Vba/Rendering/MermaidCallgraphRendererTests.cs`

The escaping is already implemented in Task 11; this task adds focused regression tests.

- [ ] **Step 1: Append tests**

```csharp
[Fact]
public void Procedure_name_with_brackets_is_escaped_in_label()
{
    var node = new CallgraphNode("M.[Bracketed Name]", "[Bracketed Name]", "M", false, false, false);
    var output = R.Render(
        new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
        new CallgraphRenderOptions(Layout: "flat"));

    Assert.DoesNotContain("[[Bracketed Name]]", output);   // would close outer label
    Assert.Contains("&#91;Bracketed Name&#93;", output);
}

[Fact]
public void Procedure_name_with_parens_is_escaped_in_handler_node()
{
    var node = new CallgraphNode("M.Foo(bar)", "Foo(bar)", "M",
        IsEventHandler: true, IsOrphan: false, IsExternal: false);
    var output = R.Render(
        new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
        new CallgraphRenderOptions(Layout: "flat"));

    // Inner parens in label escaped — outer ([ ... ]) for handler shape stays intact.
    Assert.Contains("Foo&#40;bar&#41;", output);
}

[Fact]
public void Module_name_with_space_is_mangled_in_subgraph_id()
{
    var node = new CallgraphNode("Sheet 1.P1", "P1", "Sheet 1", false, false, false);
    var output = R.Render(
        new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
        new CallgraphRenderOptions(Layout: "clustered"));

    Assert.Contains("subgraph Sheet_1", output);
}

[Fact]
public void Subgraph_open_count_matches_end_count()
{
    var nodes = new[]
    {
        new CallgraphNode("M1.A", "A", "M1", false, false, false),
        new CallgraphNode("M1.B", "B", "M1", false, false, false),
        new CallgraphNode("M2.C", "C", "M2", false, false, false),
        new CallgraphNode("__ext__MsgBox", "MsgBox", null, false, false, true),
    };
    var output = R.Render(
        new FilteredCallgraph(nodes, Array.Empty<CallgraphEdge>()),
        new CallgraphRenderOptions(Layout: "clustered"));

    var subgraphCount = System.Text.RegularExpressions.Regex.Matches(output, @"^\s*subgraph\b", System.Text.RegularExpressions.RegexOptions.Multiline).Count;
    var endCount = System.Text.RegularExpressions.Regex.Matches(output, @"^\s*end\s*$", System.Text.RegularExpressions.RegexOptions.Multiline).Count;
    Assert.Equal(subgraphCount, endCount);
    Assert.Equal(2, subgraphCount);   // M1, M2 — no subgraph for the external
}
```

- [ ] **Step 2: Run — should pass**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~MermaidCallgraphRendererTests"
```

If any fails, the implementation needs an escape-table tweak — fix and retest.

- [ ] **Step 3: Commit**

```bash
git add tests/mcpOffice.Tests/Excel/Vba/Rendering/MermaidCallgraphRendererTests.cs
git commit -m "test: Mermaid renderer — escaping and subgraph balance regressions"
```

---

### Task 13: `DotCallgraphRenderer`

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/Rendering/DotCallgraphRenderer.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/Rendering/DotCallgraphRendererTests.cs`

- [ ] **Step 1: Write failing tests**

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba.Rendering;

namespace McpOffice.Tests.Excel.Vba.Rendering;

public class DotCallgraphRendererTests
{
    private static readonly DotCallgraphRenderer R = new();

    [Fact]
    public void Empty_graph_emits_digraph_header_and_braces()
    {
        var output = R.Render(
            new FilteredCallgraph(Array.Empty<CallgraphNode>(), Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions());

        Assert.StartsWith("digraph G {", output);
        Assert.EndsWith("}\n", output);
    }

    [Fact]
    public void Single_node_clustered_wraps_in_subgraph_cluster()
    {
        var node = new CallgraphNode("M.P1", "P1", "M", false, true, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "clustered"));

        Assert.Contains("subgraph cluster_M", output);
        Assert.Contains("\"M.P1\"", output);
        Assert.Contains("label=\"P1\"", output);
    }

    [Fact]
    public void Flat_uses_FQN_label()
    {
        var node = new CallgraphNode("M.P1", "P1", "M", false, false, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.DoesNotContain("subgraph", output);
        Assert.Contains("label=\"M.P1\"", output);
    }

    [Fact]
    public void Resolved_edge_solid_unresolved_dashed()
    {
        var p1 = new CallgraphNode("M.P1", "P1", "M", false, false, false);
        var p2 = new CallgraphNode("M.P2", "P2", "M", false, false, false);
        var ext = new CallgraphNode("__ext__MsgBox", "MsgBox", null, false, false, true);

        var output = R.Render(
            new FilteredCallgraph(
                new[] { p1, p2, ext },
                new[]
                {
                    new CallgraphEdge("M.P1", "M.P2", Resolved: true),
                    new CallgraphEdge("M.P1", "__ext__MsgBox", Resolved: false),
                }),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Matches("\"M.P1\"\\s*->\\s*\"M.P2\"", output);
        Assert.Contains("style=\"dashed\"", output);
    }

    [Fact]
    public void Handler_node_uses_oval_shape()
    {
        var node = new CallgraphNode("M.Open", "Open", "M",
            IsEventHandler: true, IsOrphan: false, IsExternal: false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Contains("shape=oval", output);
    }

    [Fact]
    public void External_node_styled_dashed()
    {
        var ext = new CallgraphNode("__ext__MsgBox", "MsgBox", null, false, false, true);
        var output = R.Render(
            new FilteredCallgraph(new[] { ext }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Contains("style=\"dashed,filled\"", output);
        Assert.Contains("fillcolor=\"#f5f5f5\"", output);
    }

    [Fact]
    public void Procedure_id_with_quote_is_escaped()
    {
        var node = new CallgraphNode("M.It\"s", "It\"s", "M", false, false, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        // No unescaped " inside an identifier or label.
        Assert.Contains("\\\"", output);
    }

    [Fact]
    public void Brace_balance_holds()
    {
        var nodes = new[]
        {
            new CallgraphNode("M1.A", "A", "M1", false, false, false),
            new CallgraphNode("M2.B", "B", "M2", false, false, false),
        };
        var output = R.Render(
            new FilteredCallgraph(nodes, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "clustered"));

        Assert.Equal(output.Count(c => c == '{'), output.Count(c => c == '}'));
    }
}
```

- [ ] **Step 2: Run — they fail**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~DotCallgraphRendererTests"
```

- [ ] **Step 3: Implement**

`src/mcpOffice/Services/Excel/Vba/Rendering/DotCallgraphRenderer.cs`:

```csharp
using System.Text;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba.Rendering;

public sealed class DotCallgraphRenderer : ICallgraphRenderer
{
    public string Render(FilteredCallgraph graph, CallgraphRenderOptions options)
    {
        var sb = new StringBuilder();
        sb.AppendLine("digraph G {");
        sb.AppendLine("  rankdir=TB;");
        sb.AppendLine("  node [shape=box];");

        if (options.Layout == "clustered")
            EmitClustered(sb, graph);
        else
            EmitFlat(sb, graph);

        EmitEdges(sb, graph);
        sb.AppendLine("}");
        return sb.ToString();
    }

    private static void EmitClustered(StringBuilder sb, FilteredCallgraph graph)
    {
        var grouped = graph.Nodes
            .Where(n => !n.IsExternal)
            .GroupBy(n => n.Module!)
            .OrderBy(g => g.Key, StringComparer.Ordinal);

        foreach (var group in grouped)
        {
            var clusterId = "cluster_" + Mangle(group.Key);
            sb.Append("  subgraph ").Append(clusterId).AppendLine(" {");
            sb.Append("    label=").Append(Quote(group.Key)).AppendLine(";");
            foreach (var node in group)
            {
                sb.Append("    ");
                EmitNode(sb, node, useFqnLabel: false);
                sb.AppendLine();
            }
            sb.AppendLine("  }");
        }

        foreach (var ext in graph.Nodes.Where(n => n.IsExternal))
        {
            sb.Append("  ");
            EmitNode(sb, ext, useFqnLabel: false);
            sb.AppendLine();
        }
    }

    private static void EmitFlat(StringBuilder sb, FilteredCallgraph graph)
    {
        foreach (var node in graph.Nodes)
        {
            sb.Append("  ");
            EmitNode(sb, node, useFqnLabel: !node.IsExternal);
            sb.AppendLine();
        }
    }

    private static void EmitNode(StringBuilder sb, CallgraphNode node, bool useFqnLabel)
    {
        var id = Quote(node.Id);
        var label = useFqnLabel ? node.Id : node.Label;

        var attrs = new List<string> { $"label={Quote(label)}" };

        if (node.IsExternal)
        {
            attrs.Add("shape=box");
            attrs.Add("style=\"dashed,filled\"");
            attrs.Add("fillcolor=\"#f5f5f5\"");
        }
        else if (node.IsEventHandler)
        {
            attrs.Add("shape=oval");
            attrs.Add("style=\"filled\"");
            attrs.Add("fillcolor=\"#e1f5ff\"");
        }
        else if (node.IsOrphan)
        {
            attrs.Add("style=\"dashed\"");
        }

        sb.Append(id).Append(" [").Append(string.Join(", ", attrs)).Append("];");
    }

    private static void EmitEdges(StringBuilder sb, FilteredCallgraph graph)
    {
        foreach (var e in graph.Edges)
        {
            sb.Append("  ").Append(Quote(e.FromId)).Append(" -> ").Append(Quote(e.ToId));
            if (!e.Resolved)
                sb.Append(" [style=\"dashed\"]");
            sb.AppendLine(";");
        }
    }

    private static string Mangle(string s)
    {
        var chars = s.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray();
        return new string(chars);
    }

    private static string Quote(string s) => "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
```

- [ ] **Step 4: Run — all green**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~DotCallgraphRendererTests"
```

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/Rendering/DotCallgraphRenderer.cs tests/mcpOffice.Tests/Excel/Vba/Rendering/DotCallgraphRendererTests.cs
git commit -m "feat: DotCallgraphRenderer with clusters, flat, styling, escaping"
```

---

# Phase 5 — Service wiring

### Task 14: `IExcelWorkbookService.RenderVbaCallgraph` + impl

**Files:**
- Modify: `src/mcpOffice/Services/Excel/IExcelWorkbookService.cs`
- Modify: `src/mcpOffice/Services/Excel/ExcelWorkbookService.cs`

- [ ] **Step 1: Extend the interface**

Add to `IExcelWorkbookService`:

```csharp
string RenderVbaCallgraph(
    string path,
    string format,
    string? moduleName,
    string? procedureName,
    int depth,
    string direction,
    string layout,
    int maxNodes);
```

- [ ] **Step 2: Implement on `ExcelWorkbookService`**

Add the method alongside `AnalyzeVba` (use the same `try / catch (Exception ex) when (ex is not McpException)` wrapper):

```csharp
public string RenderVbaCallgraph(
    string path,
    string format,
    string? moduleName,
    string? procedureName,
    int depth,
    string direction,
    string layout,
    int maxNodes)
{
    PathGuard.RequireExists(path);

    ICallgraphRenderer renderer = format switch
    {
        "mermaid" => new MermaidCallgraphRenderer(),
        "dot" => new DotCallgraphRenderer(),
        _ => throw ToolError.InvalidRenderOption(
            "format", format, "Use one of mermaid, dot."),
    };

    if (layout != "clustered" && layout != "flat")
        throw ToolError.InvalidRenderOption(
            "layout", layout, "Use one of clustered, flat.");

    try
    {
        var project = new VbaProjectReader().Read(path);
        var analysis = VbaSourceAnalyzer.Analyze(
            project, includeProcedures: true, includeCallGraph: true, includeReferences: false);

        if (!analysis.HasVbaProject)
        {
            return renderer.Render(
                new FilteredCallgraph(Array.Empty<CallgraphNode>(), Array.Empty<CallgraphEdge>()),
                new CallgraphRenderOptions(layout));
        }

        var filtered = VbaCallgraphFilter.Apply(analysis,
            new CallgraphFilterOptions(
                ModuleName: moduleName,
                ProcedureName: procedureName,
                Depth: depth,
                Direction: direction,
                MaxNodes: maxNodes));

        return renderer.Render(filtered, new CallgraphRenderOptions(layout));
    }
    catch (Exception ex) when (ex is not McpException)
    {
        throw ToolError.ParseError(path, ex.Message);
    }
}
```

Required `using` additions at top of `ExcelWorkbookService.cs`:

```csharp
using McpOffice.Services.Excel.Vba.Rendering;
```

- [ ] **Step 3: Build**

```bash
dotnet build --nologo
```

Expected: 0 errors. (No new tests yet — covered by Phase 6 integration test.)

- [ ] **Step 4: Commit**

```bash
git add src/mcpOffice/Services/Excel/IExcelWorkbookService.cs src/mcpOffice/Services/Excel/ExcelWorkbookService.cs
git commit -m "feat: ExcelWorkbookService.RenderVbaCallgraph wires filter + renderer"
```

---

### Task 15: Tool registration on `ExcelTools`

**Files:**
- Modify: `src/mcpOffice/Tools/ExcelTools.cs`
- Modify: `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`

- [ ] **Step 1: Update the expected catalog test**

In `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`, add `"excel_render_vba_callgraph"` to the `expected` array (alphabetical placement: between `excel_read_sheet` and `Ping`).

- [ ] **Step 2: Run — fails**

```bash
dotnet test tests/mcpOffice.Tests.Integration --nologo --filter "FullyQualifiedName~ToolSurfaceTests"
```

Expected: `Exposes_initial_tool_catalog` fails — server doesn't expose the new tool yet.

- [ ] **Step 3: Add the tool method**

Append to `ExcelTools` (after `ExcelAnalyzeVba`):

```csharp
[McpServerTool(Name = "excel_render_vba_callgraph")]
[Description("Renders the VBA call graph as Mermaid (default) or DOT for visual inspection. Layered on excel_analyze_vba. Use moduleName / procedureName / depth / direction to narrow on large workbooks; without filters, large workbooks throw graph_too_large. Returns the rendered string directly — no JSON wrapper.")]
public static object ExcelRenderVbaCallgraph(
    [Description("Absolute path to the .xlsm/.xlsb workbook")] string path,
    [Description("Output format: 'mermaid' (default, renders inline in Markdown) or 'dot' (Graphviz).")] string format = "mermaid",
    [Description("Optional case-insensitive module name to scope the graph to a single module's neighbourhood.")] string? moduleName = null,
    [Description("Optional case-insensitive focal procedure name within moduleName. Requires moduleName.")] string? procedureName = null,
    [Description("BFS hops out from the focal procedure. Used only with procedureName. Default 2.")] int depth = 2,
    [Description("BFS direction: 'callees', 'callers', or 'both'. Used only with procedureName. Default 'both'.")] string direction = "both",
    [Description("Layout: 'clustered' (subgraph per module, default) or 'flat'.")] string layout = "clustered",
    [Description("Hard cap on rendered node count. Throws graph_too_large past this. Default 300.")] int maxNodes = 300)
    => Service.RenderVbaCallgraph(path, format, moduleName, procedureName, depth, direction, layout, maxNodes);
```

- [ ] **Step 4: Run — passes**

```bash
dotnet test tests/mcpOffice.Tests.Integration --nologo --filter "FullyQualifiedName~ToolSurfaceTests"
```

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Tools/ExcelTools.cs tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs
git commit -m "feat: register excel_render_vba_callgraph MCP tool"
```

---

# Phase 6 — Integration + benchmark

### Task 16: Stdio integration test

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs`

- [ ] **Step 1: Append the test**

```csharp
[Fact]
public async Task Render_vba_callgraph_via_stdio_returns_mermaid()
{
    var fixture = ResolveFixturePath("sample-with-macros.xlsm");
    if (!File.Exists(fixture)) return;

    await using var harness = await ServerHarness.StartAsync();
    var result = await harness.Client.CallToolAsync(
        "excel_render_vba_callgraph",
        new Dictionary<string, object?>
        {
            ["path"] = fixture,
            ["format"] = "mermaid",
            ["layout"] = "flat"
        });

    var text = result.Content.OfType<TextContentBlock>().Single().Text;
    Assert.NotEmpty(text);
    // The MCP serialiser may wrap the string as JSON; either way the rendered
    // payload starts with "flowchart" — assert the substring.
    Assert.Contains("flowchart TD", text);
}

[Fact]
public async Task Render_vba_callgraph_returns_empty_flowchart_for_xlsx_without_macros()
{
    var path = TempPath(".xlsx");
    try
    {
        using (var workbook = new Workbook())
        {
            workbook.Worksheets[0].Cells["A1"].Value = "x";
            workbook.SaveDocument(path, SpreadsheetFormat.Xlsx);
        }

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_render_vba_callgraph",
            new Dictionary<string, object?> { ["path"] = path });

        var text = result.Content.OfType<TextContentBlock>().Single().Text;
        Assert.Contains("flowchart TD", text);
        // No subgraph, no edges — just the header and classDefs.
        Assert.DoesNotContain("subgraph", text);
        Assert.DoesNotContain("-->", text);
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
```

- [ ] **Step 2: Run — passes**

```bash
dotnet test tests/mcpOffice.Tests.Integration --nologo
```

- [ ] **Step 3: Commit**

```bash
git add tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs
git commit -m "test: stdio integration for excel_render_vba_callgraph"
```

---

### Task 17: Air.xlsm gated benchmark

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Vba/AirSampleRenderTests.cs`

- [ ] **Step 1: Add the gated benchmark**

```csharp
using System.Diagnostics;
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

public class AirSampleRenderTests
{
    private const string SamplePath = @"C:\Projects\mcpOffice-samples\Air.xlsm";

    [Fact]
    public void Whole_workbook_render_throws_graph_too_large()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var act = () => svc.RenderVbaCallgraph(
            SamplePath,
            format: "mermaid",
            moduleName: null,
            procedureName: null,
            depth: 2,
            direction: "both",
            layout: "clustered",
            maxNodes: 300);

        var ex = Assert.Throws<ModelContextProtocol.McpException>(act);
        Assert.Contains("graph_too_large", ex.Message);
    }

    [Fact]
    public void Single_module_render_succeeds_under_max_nodes()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        // Air.xlsm's largest module by analyzer is "modAirData" — but if name drift bites,
        // any module that successfully parses works for the contract assertion.
        // Pick one that exists: read the analyzer first.
        var analysis = svc.AnalyzeVba(SamplePath, true, true, false);
        var probe = analysis.Modules!.First(m => m.Parsed && m.Procedures.Count > 0);

        var output = svc.RenderVbaCallgraph(
            SamplePath,
            format: "mermaid",
            moduleName: probe.Name,
            procedureName: null,
            depth: 2,
            direction: "both",
            layout: "clustered",
            maxNodes: 300);

        Assert.StartsWith("flowchart TD", output);
        Assert.Contains("subgraph", output);
    }

    [Fact]
    public void Pipeline_completes_under_500ms()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var analysis = svc.AnalyzeVba(SamplePath, true, true, false);
        var probe = analysis.Modules!.First(m => m.Parsed && m.Procedures.Count > 0);

        var sw = Stopwatch.StartNew();
        _ = svc.RenderVbaCallgraph(
            SamplePath,
            format: "mermaid",
            moduleName: probe.Name,
            procedureName: probe.Procedures[0].Name,
            depth: 1,
            direction: "both",
            layout: "clustered",
            maxNodes: 300);
        sw.Stop();

        Assert.True(sw.ElapsedMilliseconds < 500,
            $"render pipeline took {sw.ElapsedMilliseconds}ms, expected < 500ms");
    }
}
```

- [ ] **Step 2: Run on a machine with the sample**

```bash
dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~AirSampleRenderTests"
```

Expected: 3 passing if `C:\Projects\mcpOffice-samples\Air.xlsm` exists; 3 silently no-op (passing) otherwise.

- [ ] **Step 3: Commit**

```bash
git add tests/mcpOffice.Tests/Excel/Vba/AirSampleRenderTests.cs
git commit -m "test: gated Air.xlsm render benchmark (cap, module render, wall time)"
```

---

# Phase 7 — Final verification + PR

### Task 18: Full verification

- [ ] **Step 1: Clean build**

```bash
dotnet build --nologo -c Release
```

Expected: 0 warnings, 0 errors.

- [ ] **Step 2: Full test run**

```bash
dotnet test --nologo -c Release
```

Expected: every test passing. Tool count visible in `ToolSurfaceTests` should be 25 (24 + new tool).

- [ ] **Step 3: Manual smoke (optional but recommended)**

Restart Claude Code so it picks up the new tool, then call:

```
mcp__office__excel_render_vba_callgraph(
    path="C:\\Projects\\mcpOffice-samples\\Air.xlsm",
    moduleName="modAirData",     // or whatever Air.xlsm exposes
    layout="clustered"
)
```

Expected: a Mermaid flowchart that renders inline in chat. Per global CLAUDE.md, build green ≠ feature works — confirm visually that the picture is sensible (event handlers stand out, modules are clustered, externals are dashed).

### Task 19: Open PR

- [ ] **Step 1: Push branch**

```bash
git push -u origin feat/render-vba-callgraph
```

- [ ] **Step 2: Open PR**

```bash
gh pr create --title "feat: excel_render_vba_callgraph — Mermaid/DOT call-graph renderer" --body "$(cat <<'EOF'
## Summary
- New MCP tool `excel_render_vba_callgraph` emits the VBA call graph as Mermaid (default) or DOT.
- Layered on `excel_analyze_vba` — same parse pass, new filter + renderer pipeline.
- Three filter modes: whole workbook, `moduleName` direct neighbourhood, focal `procedureName` + depth + direction.
- Module-clustered layout default; flat as an option.
- Visual conventions for event handlers, orphans, and external (unresolved) callees.
- 300-node hard cap with `graph_too_large`; no silent truncation.

Design: docs/plans/2026-05-03-mcpoffice-excel-render-vba-callgraph-design.md
Plan: docs/plans/2026-05-03-mcpoffice-excel-render-vba-callgraph-plan.md

## Test plan
- [x] dotnet build clean (0 warnings, 0 errors)
- [x] dotnet test green (filter, renderer, integration, gated benchmark)
- [x] Tool count 24 → 25 in ToolSurfaceTests
- [x] Manual smoke against Air.xlsm — Mermaid renders inline, clusters per module, externals dashed

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

- [ ] **Step 3: Update SESSION_HANDOFF.md and TODO.md after merge**

After squash-merging:

```bash
git checkout main
git pull --ff-only
```

Then:
- `SESSION_HANDOFF.md` — bump latest commit, tool surface (25), next-up section
- `TODO.md` — mark "Dependency graph rendering" as done; conversion-hints and coupling-score remain

---

## What this plan deliberately does NOT do

- **No `includeExternal` toggle.** Externals are always rendered. Add the toggle only if a real workbook makes it noisy.
- **No SVG/PNG rasterisation.** The agent gets text; downstream tools render.
- **No DiagramControl / Excalidraw output.** Out of scope per design.
- **No conversion hints / coupling score.** v3 / v4.
- **No caching across calls.** Each invocation re-parses. Cheap relative to the rest of the work.

## Risks called out

1. **Mermaid escape coverage.** Real-world workbook procedure names (Dutch, with `[brackets]`, with `_` continuations parsed as part of the name) may surface escape edge cases. Tests in Task 12 cover the obvious ones; if a real workbook breaks the renderer, the fix is one new test + an additional escape entry.
2. **Module name collisions in DOT cluster IDs.** `Mangle("Sheet 1")` and `Mangle("Sheet_1")` both produce `Sheet_1`. Unlikely in real workbooks but worth noting; if it bites, switch to a `cluster_<index>` numbering scheme.
3. **Air.xlsm test brittleness.** `AirSampleRenderTests` picks the first module with procedures by analyzer order — if the underlying ordering shifts, the wall-time test still passes but the module under test changes silently. Acceptable for a benchmark; not a correctness gate.
