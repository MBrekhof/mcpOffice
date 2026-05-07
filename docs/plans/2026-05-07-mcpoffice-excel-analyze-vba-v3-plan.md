# `excel_suggest_vba_conversion` (analyzer v3) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship `excel_suggest_vba_conversion` as the 26th MCP tool — a conversion-hints layer that consumes v1's structural model and emits per-procedure migration hints (multi-axis tagging) plus workbook-wide module coupling, with an optional `targetParadigm` overlay producing structured C# emission targets.

**Architecture:** New service entry point `ExcelWorkbookService.SuggestVbaConversion` that calls `VbaSourceAnalyzer.Analyze` *unfiltered* (the coupling block needs the whole-workbook graph), then runs a new `VbaConversionHintBuilder` that decomposes into three pure sub-components (`AxisClassifier`, `CouplingComputer`, `ParadigmOverlayApplier`). `moduleName` filtering happens in the builder, after the analyzer returns. No new VBA parsing — every output is derived from existing v1 records.

**Tech Stack:** .NET 9, C#, MCP SDK 1.2.0, xUnit, ModelContextProtocol C# SDK, DevExpress.Document.Processor (transitive — through v1's analyzer path).

**Reference design:** `docs/plans/2026-05-07-mcpoffice-excel-analyze-vba-v3-design.md` — single source of truth for tool surface, axes, coupling rules, paradigm matrix, error codes. Read it before starting.

---

## Conventions used in this plan

- All paths are relative to `C:\Projects\mcpOffice\` (the repo root).
- Working branch: `feat/excel-vba-conversion-hints-v3` (already exists; design doc landed at `3d2dd39`).
- "Run build" = `dotnet build --nologo`. "Run tests" = `dotnet test --nologo`.
- After every task: `dotnet build` is green AND every test passes. If either fails, stop and fix before moving on (per superpowers:verification-before-completion).
- Conventional Commits (`feat:`, `test:`, `chore:`, `docs:`).
- The analyzer's records (`ExcelVbaAnalysis`, `ExcelVbaProcedure`, `ExcelVbaCallEdge`, `ExcelVbaObjectModelRef`, `ExcelVbaDependency`, `ExcelVbaModuleAnalysis`, `ExcelVbaSiteRef`) are read-only inputs — the plan never mutates v1.
- Tests against the analyzer's records use the same synthetic-record helper pattern as `VbaCallgraphFilterTests` (no real `.xlsm` parsing in axis/coupling/overlay unit tests).

## File structure overview

**New under `src/mcpOffice/`:**
- `Models/ConversionHints.cs` — top-level result record.
- `Models/ProcedureHint.cs` — per-procedure entry with axes + suggestion.
- `Models/ProcedureAxes.cs` — the four-axis classification.
- `Models/CSharpSuggestion.cs` — emission target (paradigm overlay).
- `Models/ModuleCoupling.cs` — Ca/Ce/I per module.
- `Models/CouplingPair.cs` — directional pairwise edge weight.
- `Models/ConversionHintsSummary.cs` — small summary block.
- `Services/Excel/Vba/AxisClassifier.cs` — pure function: procedure + analysis context → axes.
- `Services/Excel/Vba/CouplingComputer.cs` — pure function: call graph + module list → coupling block.
- `Services/Excel/Vba/ParadigmOverlayApplier.cs` — pure function: axes + identity + paradigm → suggestion.
- `Services/Excel/Vba/VbaConversionHintBuilder.cs` — orchestrator combining the three above.

**Modified under `src/mcpOffice/`:**
- `ErrorCode.cs` — add `UnsupportedParadigm`.
- `ToolError.cs` — add `UnsupportedParadigm` helper.
- `Services/Excel/IExcelWorkbookService.cs` — add `SuggestVbaConversion`.
- `Services/Excel/ExcelWorkbookService.cs` — implement `SuggestVbaConversion`.
- `Tools/ExcelTools.cs` — add `[McpServerTool(Name="excel_suggest_vba_conversion")]`.

**New under `tests/mcpOffice.Tests/Excel/Vba/`:**
- `AxisClassifierTests.cs`
- `CouplingComputerTests.cs`
- `ParadigmOverlayApplierTests.cs`
- `VbaConversionHintBuilderTests.cs`
- `SyntheticConversionHintsTests.cs`
- `AirSampleConversionHintsTests.cs`

**Modified under `tests/`:**
- `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs` — add new tool name (25 → 26).
- `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs` — one round-trip happy-path stdio test.

---

# Phase 1 — Foundations

### Task 1: Add `unsupported_paradigm` error code + helper

**Files:**
- Modify: `src/mcpOffice/ErrorCode.cs`
- Modify: `src/mcpOffice/ToolError.cs`
- Test: `tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs` (existing — append a fact)

- [ ] **Step 1: Write the failing test**

Append to `tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs`:

```csharp
[Fact]
public void UnsupportedParadigm_throws_McpException_with_code_in_message()
{
    var act = () => throw ToolError.UnsupportedParadigm("blazor", new[] { "classLibrary", "workerService" });
    var ex = Assert.Throws<ModelContextProtocol.McpException>(act);
    Assert.Contains("[unsupported_paradigm]", ex.Message);
    Assert.Contains("blazor", ex.Message);
    Assert.Contains("classLibrary", ex.Message);
}
```

- [ ] **Step 2: Run test to verify it fails**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~UnsupportedParadigm --nologo
```

Expected: FAIL — `'ToolError' does not contain a definition for 'UnsupportedParadigm'`.

- [ ] **Step 3: Add the error code**

Add to `src/mcpOffice/ErrorCode.cs` after `InvalidRenderOption`:

```csharp
    public const string UnsupportedParadigm = "unsupported_paradigm";
```

- [ ] **Step 4: Add the helper**

Add to `src/mcpOffice/ToolError.cs` after `InvalidRenderOption`:

```csharp
    public static Exception UnsupportedParadigm(string paradigm, IEnumerable<string> supported) =>
        Throw(ErrorCode.UnsupportedParadigm,
            $"Unsupported targetParadigm: {paradigm}. Supported values: {string.Join(", ", supported)}");
```

- [ ] **Step 5: Run test to verify it passes**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~UnsupportedParadigm --nologo
```

Expected: PASS.

- [ ] **Step 6: Run the full build**

```
dotnet build --nologo
```

Expected: 0 warnings, 0 errors.

- [ ] **Step 7: Commit**

```bash
git add src/mcpOffice/ErrorCode.cs src/mcpOffice/ToolError.cs tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs
git commit -m "feat: add unsupported_paradigm error code"
```

---

### Task 2: Add ConversionHints DTOs

**Files:**
- Create: `src/mcpOffice/Models/ProcedureAxes.cs`
- Create: `src/mcpOffice/Models/CSharpSuggestion.cs`
- Create: `src/mcpOffice/Models/ProcedureHint.cs`
- Create: `src/mcpOffice/Models/ModuleCoupling.cs`
- Create: `src/mcpOffice/Models/CouplingPair.cs`
- Create: `src/mcpOffice/Models/ConversionHintsSummary.cs`
- Create: `src/mcpOffice/Models/ConversionHints.cs`

This is a pure type-definition task — no behaviour, no test. The shapes must match the design doc's Output schema.

- [ ] **Step 1: Create `ProcedureAxes.cs`**

```csharp
namespace McpOffice.Models;

public sealed record ProcedureAxes(
    string Trigger,
    string Purity,
    string? Shape,
    IReadOnlyList<string> Dependencies);
```

- [ ] **Step 2: Create `CSharpSuggestion.cs`**

```csharp
namespace McpOffice.Models;

public sealed record CSharpSuggestion(
    string TargetType,
    string SuggestedClassName,
    string SuggestedMethodName,
    string? Lifetime,
    bool IsPublic,
    IReadOnlyList<string> Blockers);
```

- [ ] **Step 3: Create `ProcedureHint.cs`**

```csharp
namespace McpOffice.Models;

public sealed record ProcedureHint(
    string Module,
    string ProcedureName,
    string Kind,
    bool IsEventHandler,
    int ParamCount,
    int CallerCount,
    int CalleeCount,
    ProcedureAxes Axes,
    string Rationale,
    CSharpSuggestion? CsharpSuggestion);
```

- [ ] **Step 4: Create `ModuleCoupling.cs`**

```csharp
namespace McpOffice.Models;

public sealed record ModuleCoupling(
    string Module,
    int Ca,
    int Ce,
    double Instability,
    int InternalEdges);
```

- [ ] **Step 5: Create `CouplingPair.cs`**

```csharp
namespace McpOffice.Models;

public sealed record CouplingPair(string From, string To, int EdgeCount);
```

- [ ] **Step 6: Create `ConversionHintsSummary.cs`**

```csharp
namespace McpOffice.Models;

public sealed record ConversionHintsSummary(
    int TotalProcedures,
    int HintedProcedures,
    int ModuleCount,
    string? TargetParadigm,
    long WallTimeMs);
```

- [ ] **Step 7: Create `ConversionHints.cs`**

```csharp
namespace McpOffice.Models;

public sealed record ConversionHints(
    ConversionHintsSummary Summary,
    IReadOnlyList<ProcedureHint> ProcedureHints,
    IReadOnlyList<ModuleCoupling> ModuleCoupling,
    IReadOnlyList<CouplingPair> CouplingPairs);
```

- [ ] **Step 8: Run build**

```
dotnet build --nologo
```

Expected: 0 warnings, 0 errors.

- [ ] **Step 9: Commit**

```bash
git add src/mcpOffice/Models/ProcedureAxes.cs src/mcpOffice/Models/CSharpSuggestion.cs src/mcpOffice/Models/ProcedureHint.cs src/mcpOffice/Models/ModuleCoupling.cs src/mcpOffice/Models/CouplingPair.cs src/mcpOffice/Models/ConversionHintsSummary.cs src/mcpOffice/Models/ConversionHints.cs
git commit -m "feat: ConversionHints DTOs for analyzer v3"
```

---

# Phase 2 — AxisClassifier

The classifier is a pure static helper. To keep tests self-contained, every test builds a synthetic `ExcelVbaAnalysis` shape via a private helper, mirroring the pattern used by `VbaCallgraphFilterTests`.

### Task 3: AxisClassifier skeleton + trigger axis

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs`

The classifier exposes a single method `Classify(procedure, moduleKind, callGraph, references) -> ProcedureAxes`. The four axes are computed independently inside `Classify`; this task implements only `trigger`. Other axes return placeholder values that the next tasks will replace.

- [ ] **Step 1: Write the failing tests for the trigger axis**

Create `tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs`:

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class AxisClassifierTests
{
    private static ExcelVbaProcedure Proc(
        string module,
        string name,
        bool isEventHandler = false,
        string? scope = null) =>
        new(
            Name: name,
            FullyQualifiedName: $"{module}.{name}",
            Kind: "Sub",
            Scope: scope,
            Parameters: Array.Empty<ExcelVbaParameter>(),
            ReturnType: null,
            LineStart: 1,
            LineEnd: 2,
            IsEventHandler: isEventHandler,
            EventTarget: null);

    private static ExcelVbaCallEdge Edge(string from, string to, bool resolved = true)
    {
        var fromParts = from.Split('.');
        return new ExcelVbaCallEdge(
            From: from,
            To: to,
            Resolved: resolved,
            Site: new ExcelVbaSiteRef(fromParts[0], fromParts[1], 1));
    }

    [Fact]
    public void Trigger_eventHandler_when_procedure_IsEventHandler()
    {
        var proc = Proc("Sheet1", "Worksheet_Change", isEventHandler: true);
        var axes = AxisClassifier.Classify(
            proc, moduleKind: "documentModule",
            callGraph: Array.Empty<ExcelVbaCallEdge>(),
            objectModel: Array.Empty<ExcelVbaObjectModelRef>(),
            dependencies: Array.Empty<ExcelVbaDependency>());
        Assert.Equal("eventHandler", axes.Trigger);
    }

    [Fact]
    public void Trigger_macroEntryPoint_when_public_no_callers_standard_module()
    {
        var proc = Proc("Module1", "Main", scope: "Public");
        var axes = AxisClassifier.Classify(
            proc, moduleKind: "standard",
            callGraph: Array.Empty<ExcelVbaCallEdge>(),
            objectModel: Array.Empty<ExcelVbaObjectModelRef>(),
            dependencies: Array.Empty<ExcelVbaDependency>());
        Assert.Equal("macroEntryPoint", axes.Trigger);
    }

    [Fact]
    public void Trigger_macroEntryPoint_when_scope_null_no_callers_standard_module()
    {
        // VBA default scope is Public when omitted.
        var proc = Proc("Module1", "Main", scope: null);
        var axes = AxisClassifier.Classify(
            proc, moduleKind: "standard",
            callGraph: Array.Empty<ExcelVbaCallEdge>(),
            objectModel: Array.Empty<ExcelVbaObjectModelRef>(),
            dependencies: Array.Empty<ExcelVbaDependency>());
        Assert.Equal("macroEntryPoint", axes.Trigger);
    }

    [Fact]
    public void Trigger_calledOnly_when_private_orphan()
    {
        var proc = Proc("Module1", "Helper", scope: "Private");
        var axes = AxisClassifier.Classify(
            proc, moduleKind: "standard",
            callGraph: Array.Empty<ExcelVbaCallEdge>(),
            objectModel: Array.Empty<ExcelVbaObjectModelRef>(),
            dependencies: Array.Empty<ExcelVbaDependency>());
        Assert.Equal("calledOnly", axes.Trigger);
    }

    [Fact]
    public void Trigger_calledOnly_when_has_callers()
    {
        var proc = Proc("Module1", "Helper", scope: "Public");
        var edges = new[] { Edge("Module1.Main", "Module1.Helper") };
        var axes = AxisClassifier.Classify(
            proc, moduleKind: "standard",
            callGraph: edges,
            objectModel: Array.Empty<ExcelVbaObjectModelRef>(),
            dependencies: Array.Empty<ExcelVbaDependency>());
        Assert.Equal("calledOnly", axes.Trigger);
    }

    [Fact]
    public void Trigger_calledOnly_when_documentModule_kind()
    {
        // documentModule excludes macroEntryPoint even with no callers.
        var proc = Proc("Sheet1", "DoStuff", scope: "Public");
        var axes = AxisClassifier.Classify(
            proc, moduleKind: "documentModule",
            callGraph: Array.Empty<ExcelVbaCallEdge>(),
            objectModel: Array.Empty<ExcelVbaObjectModelRef>(),
            dependencies: Array.Empty<ExcelVbaDependency>());
        Assert.Equal("calledOnly", axes.Trigger);
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AxisClassifierTests --nologo
```

Expected: FAIL — `AxisClassifier` doesn't exist.

- [ ] **Step 3: Implement the skeleton with the trigger axis**

Create `src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs`:

```csharp
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class AxisClassifier
{
    public static ProcedureAxes Classify(
        ExcelVbaProcedure proc,
        string moduleKind,
        IReadOnlyList<ExcelVbaCallEdge> callGraph,
        IReadOnlyList<ExcelVbaObjectModelRef> objectModel,
        IReadOnlyList<ExcelVbaDependency> dependencies)
    {
        var trigger = ClassifyTrigger(proc, moduleKind, callGraph);
        var purity = "pure";                 // implemented in Task 4
        string? shape = null;                // implemented in Task 5
        IReadOnlyList<string> deps = Array.Empty<string>(); // implemented in Task 6
        return new ProcedureAxes(trigger, purity, shape, deps);
    }

    private static string ClassifyTrigger(
        ExcelVbaProcedure proc,
        string moduleKind,
        IReadOnlyList<ExcelVbaCallEdge> callGraph)
    {
        if (proc.IsEventHandler) return "eventHandler";

        bool isPrivate = string.Equals(proc.Scope, "Private", StringComparison.OrdinalIgnoreCase);
        bool isPublic = !isPrivate;
        bool isDocumentModule = string.Equals(moduleKind, "documentModule", StringComparison.OrdinalIgnoreCase);
        bool hasCallers = callGraph.Any(e => string.Equals(e.To, proc.FullyQualifiedName, StringComparison.OrdinalIgnoreCase));

        if (isPublic && !hasCallers && !isDocumentModule) return "macroEntryPoint";
        return "calledOnly";
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AxisClassifierTests --nologo
```

Expected: PASS — 6/6 trigger tests green.

- [ ] **Step 5: Run the full test suite to confirm nothing else broke**

```
dotnet test --nologo
```

Expected: PASS overall.

- [ ] **Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs
git commit -m "feat: AxisClassifier — trigger axis"
```

---

### Task 4: AxisClassifier — purity axis

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs`

Purity uses ONLY existing v1 data — `ObjectModelReference.Mode is read|write` (the `Api` field is unused here) and `ExternalDependencyReference.Kind`. No source-code regex. The design doc's "module-scope-write detection" is deferred to a follow-up; readsState/writesState distinction comes purely from `ObjectModelRef.Mode`.

> **Note on `ExcelVbaObjectModelRef.Mode`:** The current record `ExcelVbaObjectModelRef(Module, Procedure, Line, Api, Literal)` has no `Mode` field. The design doc references it; in practice, the `VbaReferenceCollector` populates this elsewhere and we'll use `Api` patterns to infer write intent here — Apis like `Cells`, `Range`, `Value` accessed in an assignment are writes, but without `Mode` we treat any object-model touch as `readsState` for safety. **For v3 v1: anything with object-model refs but no external deps gets `readsState`.** Promote to `writesState` later when `Mode` lands. This is simpler and the design doc's "Open questions deferred to implementation" already flags this area.

So the four-way classification reduces to:
- `pure` — no ObjectModelRef AND no Dependency for this procedure.
- `readsState` — at least one ObjectModelRef AND no Dependency.
- `writesState` — *not used in v1 of v3*; will activate when `Mode` is populated.
- `sideEffectful` — at least one Dependency (any `Kind`).

- [ ] **Step 1: Append failing tests**

Append to `AxisClassifierTests.cs`:

```csharp
[Fact]
public void Purity_pure_when_no_refs_no_deps()
{
    var proc = Proc("Module1", "Pure");
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        Array.Empty<ExcelVbaObjectModelRef>(),
        Array.Empty<ExcelVbaDependency>());
    Assert.Equal("pure", axes.Purity);
}

[Fact]
public void Purity_readsState_when_object_model_present_no_deps()
{
    var proc = Proc("Module1", "Reads");
    var refs = new[]
    {
        new ExcelVbaObjectModelRef("Module1", "Reads", 5, "Worksheets", null)
    };
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        refs,
        Array.Empty<ExcelVbaDependency>());
    Assert.Equal("readsState", axes.Purity);
}

[Fact]
public void Purity_sideEffectful_when_dependency_present()
{
    var proc = Proc("Module1", "WritesFile");
    var deps = new[]
    {
        new ExcelVbaDependency("Module1", "WritesFile", 7, "filesystem", @"C:\out.txt", "write")
    };
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        Array.Empty<ExcelVbaObjectModelRef>(),
        deps);
    Assert.Equal("sideEffectful", axes.Purity);
}

[Fact]
public void Purity_sideEffectful_supersedes_object_model()
{
    var proc = Proc("Module1", "Both");
    var refs = new[] { new ExcelVbaObjectModelRef("Module1", "Both", 5, "Range", null) };
    var deps = new[] { new ExcelVbaDependency("Module1", "Both", 6, "database", "DSN=foo", "query") };
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        refs,
        deps);
    Assert.Equal("sideEffectful", axes.Purity);
}

[Fact]
public void Purity_filters_to_only_this_procedures_refs()
{
    var proc = Proc("Module1", "Pure");
    // Refs/deps belong to a different procedure — should not affect ours.
    var refs = new[] { new ExcelVbaObjectModelRef("Module1", "Other", 10, "Range", null) };
    var deps = new[] { new ExcelVbaDependency("Module1", "Other", 11, "filesystem", null, null) };
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        refs,
        deps);
    Assert.Equal("pure", axes.Purity);
}
```

- [ ] **Step 2: Run tests to verify failure**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AxisClassifierTests --nologo
```

Expected: 5 new tests fail (purity is hardcoded to `"pure"`).

- [ ] **Step 3: Implement purity classification**

Replace the `var purity = "pure";` line in `AxisClassifier.Classify` and add a helper. Final body of `Classify`:

```csharp
        var trigger = ClassifyTrigger(proc, moduleKind, callGraph);
        var purity = ClassifyPurity(proc, objectModel, dependencies);
        string? shape = null;                // implemented in Task 5
        IReadOnlyList<string> deps = Array.Empty<string>(); // implemented in Task 6
        return new ProcedureAxes(trigger, purity, shape, deps);
```

Add the helper method:

```csharp
    private static string ClassifyPurity(
        ExcelVbaProcedure proc,
        IReadOnlyList<ExcelVbaObjectModelRef> objectModel,
        IReadOnlyList<ExcelVbaDependency> dependencies)
    {
        bool hasOwnDep = dependencies.Any(d =>
            string.Equals(d.Module, ProcModule(proc), StringComparison.OrdinalIgnoreCase) &&
            string.Equals(d.Procedure, proc.Name, StringComparison.OrdinalIgnoreCase));
        if (hasOwnDep) return "sideEffectful";

        bool hasOwnObjRef = objectModel.Any(r =>
            string.Equals(r.Module, ProcModule(proc), StringComparison.OrdinalIgnoreCase) &&
            string.Equals(r.Procedure, proc.Name, StringComparison.OrdinalIgnoreCase));
        return hasOwnObjRef ? "readsState" : "pure";
    }

    private static string ProcModule(ExcelVbaProcedure proc) =>
        proc.FullyQualifiedName.Split('.', 2)[0];
```

- [ ] **Step 4: Run tests to verify pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AxisClassifierTests --nologo
```

Expected: PASS — 11/11 (6 trigger + 5 purity).

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs
git commit -m "feat: AxisClassifier — purity axis"
```

---

### Task 5: AxisClassifier — shape axis

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs`

`shape` is `leaf` when calleeCount == 0, `orchestrator` when calleeCount >= 3, omitted (null) otherwise.

- [ ] **Step 1: Append failing tests**

```csharp
[Fact]
public void Shape_leaf_when_no_callees()
{
    var proc = Proc("Module1", "P");
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        Array.Empty<ExcelVbaObjectModelRef>(),
        Array.Empty<ExcelVbaDependency>());
    Assert.Equal("leaf", axes.Shape);
}

[Fact]
public void Shape_orchestrator_when_three_callees()
{
    var proc = Proc("Module1", "P");
    var edges = new[]
    {
        Edge("Module1.P", "Module1.A"),
        Edge("Module1.P", "Module1.B"),
        Edge("Module1.P", "Module1.C")
    };
    var axes = AxisClassifier.Classify(
        proc, "standard", edges,
        Array.Empty<ExcelVbaObjectModelRef>(),
        Array.Empty<ExcelVbaDependency>());
    Assert.Equal("orchestrator", axes.Shape);
}

[Fact]
public void Shape_null_when_one_or_two_callees()
{
    var proc = Proc("Module1", "P");
    var edges = new[]
    {
        Edge("Module1.P", "Module1.A"),
        Edge("Module1.P", "Module1.B")
    };
    var axes = AxisClassifier.Classify(
        proc, "standard", edges,
        Array.Empty<ExcelVbaObjectModelRef>(),
        Array.Empty<ExcelVbaDependency>());
    Assert.Null(axes.Shape);
}

[Fact]
public void Shape_orchestrator_includes_unresolved_callees()
{
    // Unresolved edges still count as fan-out for shape purposes.
    var proc = Proc("Module1", "P");
    var edges = new[]
    {
        Edge("Module1.P", "Module1.A"),
        Edge("Module1.P", "Module1.B"),
        Edge("Module1.P", "X.Unknown", resolved: false)
    };
    var axes = AxisClassifier.Classify(
        proc, "standard", edges,
        Array.Empty<ExcelVbaObjectModelRef>(),
        Array.Empty<ExcelVbaDependency>());
    Assert.Equal("orchestrator", axes.Shape);
}
```

- [ ] **Step 2: Run tests to verify failure**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AxisClassifierTests --nologo
```

Expected: 4 new tests fail.

- [ ] **Step 3: Implement shape classification**

Replace `string? shape = null;` with `var shape = ClassifyShape(proc, callGraph);` and add the helper:

```csharp
    private static string? ClassifyShape(
        ExcelVbaProcedure proc,
        IReadOnlyList<ExcelVbaCallEdge> callGraph)
    {
        int calleeCount = callGraph.Count(e =>
            string.Equals(e.From, proc.FullyQualifiedName, StringComparison.OrdinalIgnoreCase));
        return calleeCount switch
        {
            0 => "leaf",
            >= 3 => "orchestrator",
            _ => null
        };
    }
```

- [ ] **Step 4: Run tests to verify pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AxisClassifierTests --nologo
```

Expected: PASS — 15/15.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs
git commit -m "feat: AxisClassifier — shape axis"
```

---

### Task 6: AxisClassifier — dependencies axis

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs`

The `dependencies` axis is a sorted, deduped string array drawn from:
- `excelObjectModel` — added when any `ObjectModelRef` for this proc exists.
- `ExternalDependencyReference.Kind` (one of `filesystem`, `database`, `network`, `registry`, `shell`, `automation`).

The design doc lists `automation` is collected by v1 but is not in the final dependencies set; map `automation` → `shell` to keep the closed set tight, OR pass through. **Decision for v3: pass through the v1 `Kind` verbatim — the design doc's allowed set is `{excelObjectModel, filesystem, database, network, registry, shell}`, and `automation` is what v1 currently emits for shell-out / Application.Run; rename in this layer to `shell` so the consumer sees the final set.**

- [ ] **Step 1: Append failing tests**

```csharp
[Fact]
public void Dependencies_empty_when_no_refs()
{
    var proc = Proc("Module1", "P");
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        Array.Empty<ExcelVbaObjectModelRef>(),
        Array.Empty<ExcelVbaDependency>());
    Assert.Empty(axes.Dependencies);
}

[Fact]
public void Dependencies_includes_excelObjectModel_when_object_refs_present()
{
    var proc = Proc("Module1", "P");
    var refs = new[] { new ExcelVbaObjectModelRef("Module1", "P", 5, "Range", null) };
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        refs,
        Array.Empty<ExcelVbaDependency>());
    Assert.Contains("excelObjectModel", axes.Dependencies);
}

[Fact]
public void Dependencies_dedup_and_sorted()
{
    var proc = Proc("Module1", "P");
    var refs = new[]
    {
        new ExcelVbaObjectModelRef("Module1", "P", 1, "Range", null),
        new ExcelVbaObjectModelRef("Module1", "P", 2, "Worksheets", null)
    };
    var deps = new[]
    {
        new ExcelVbaDependency("Module1", "P", 5, "filesystem", null, null),
        new ExcelVbaDependency("Module1", "P", 6, "database", null, null),
        new ExcelVbaDependency("Module1", "P", 7, "filesystem", null, null) // duplicate kind
    };
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        refs, deps);
    Assert.Equal(new[] { "database", "excelObjectModel", "filesystem" }, axes.Dependencies);
}

[Fact]
public void Dependencies_maps_automation_to_shell()
{
    var proc = Proc("Module1", "P");
    var deps = new[] { new ExcelVbaDependency("Module1", "P", 5, "automation", null, null) };
    var axes = AxisClassifier.Classify(
        proc, "standard",
        Array.Empty<ExcelVbaCallEdge>(),
        Array.Empty<ExcelVbaObjectModelRef>(),
        deps);
    Assert.Contains("shell", axes.Dependencies);
    Assert.DoesNotContain("automation", axes.Dependencies);
}
```

- [ ] **Step 2: Run tests to verify failure**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AxisClassifierTests --nologo
```

Expected: 4 new tests fail (dependencies hardcoded to empty).

- [ ] **Step 3: Implement dependencies classification**

Replace `IReadOnlyList<string> deps = Array.Empty<string>();` with `var deps = ClassifyDependencies(proc, objectModel, dependencies);` and add:

```csharp
    private static IReadOnlyList<string> ClassifyDependencies(
        ExcelVbaProcedure proc,
        IReadOnlyList<ExcelVbaObjectModelRef> objectModel,
        IReadOnlyList<ExcelVbaDependency> dependencies)
    {
        var module = ProcModule(proc);
        var set = new SortedSet<string>(StringComparer.Ordinal);

        if (objectModel.Any(r =>
                string.Equals(r.Module, module, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(r.Procedure, proc.Name, StringComparison.OrdinalIgnoreCase)))
        {
            set.Add("excelObjectModel");
        }

        foreach (var d in dependencies)
        {
            if (!string.Equals(d.Module, module, StringComparison.OrdinalIgnoreCase)) continue;
            if (!string.Equals(d.Procedure, proc.Name, StringComparison.OrdinalIgnoreCase)) continue;
            var kind = string.Equals(d.Kind, "automation", StringComparison.OrdinalIgnoreCase)
                ? "shell"
                : d.Kind;
            set.Add(kind);
        }

        return set.ToArray();
    }
```

- [ ] **Step 4: Run tests to verify pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AxisClassifierTests --nologo
```

Expected: PASS — 19/19.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/AxisClassifier.cs tests/mcpOffice.Tests/Excel/Vba/AxisClassifierTests.cs
git commit -m "feat: AxisClassifier — dependencies axis"
```

---

# Phase 3 — CouplingComputer

### Task 7: CouplingComputer — moduleCoupling

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/CouplingComputer.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/CouplingComputerTests.cs`

Computes `Ca`, `Ce`, `instability`, `internalEdges` per module, plus `couplingPairs[]`. Single pass over the call graph; whole-workbook scope only. Skips unresolved edges and dedupes by `(fromModule, fromProc, toModule, toProc)`.

This task implements the per-module half. Pairs come in Task 8.

- [ ] **Step 1: Write the failing tests**

Create `tests/mcpOffice.Tests/Excel/Vba/CouplingComputerTests.cs`:

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class CouplingComputerTests
{
    private static ExcelVbaCallEdge Edge(string from, string to, bool resolved = true) =>
        new(from, to, resolved,
            new ExcelVbaSiteRef(from.Split('.')[0], from.Split('.')[1], 1));

    private static IReadOnlyList<(string Name, string Kind)> Modules(params string[] names) =>
        names.Select(n => (n, "standard")).ToList();

    [Fact]
    public void ModuleCoupling_zeros_for_isolated_module()
    {
        var result = CouplingComputer.Compute(
            Modules("Lonely"),
            Array.Empty<ExcelVbaCallEdge>());
        var m = Assert.Single(result.Coupling);
        Assert.Equal("Lonely", m.Module);
        Assert.Equal(0, m.Ca);
        Assert.Equal(0, m.Ce);
        Assert.Equal(0.0, m.Instability);
        Assert.Equal(0, m.InternalEdges);
    }

    [Fact]
    public void ModuleCoupling_counts_external_edges_for_ca_and_ce()
    {
        // A → B (1 edge). A: Ce=1, Ca=0. B: Ca=1, Ce=0.
        var edges = new[] { Edge("A.f", "B.g") };
        var result = CouplingComputer.Compute(Modules("A", "B"), edges);

        var a = result.Coupling.Single(c => c.Module == "A");
        Assert.Equal(0, a.Ca); Assert.Equal(1, a.Ce); Assert.Equal(1.0, a.Instability);

        var b = result.Coupling.Single(c => c.Module == "B");
        Assert.Equal(1, b.Ca); Assert.Equal(0, b.Ce); Assert.Equal(0.0, b.Instability);
    }

    [Fact]
    public void ModuleCoupling_internalEdges_counts_intra_module_edges()
    {
        var edges = new[] { Edge("A.f", "A.g"), Edge("A.g", "A.h") };
        var result = CouplingComputer.Compute(Modules("A"), edges);
        var a = result.Coupling.Single();
        Assert.Equal(2, a.InternalEdges);
        Assert.Equal(0, a.Ca);
        Assert.Equal(0, a.Ce);
    }

    [Fact]
    public void ModuleCoupling_excludes_unresolved_edges()
    {
        var edges = new[] { Edge("A.f", "Unknown.x", resolved: false) };
        var result = CouplingComputer.Compute(Modules("A", "B"), edges);
        Assert.All(result.Coupling, m => Assert.Equal(0, m.Ce));
    }

    [Fact]
    public void ModuleCoupling_dedupes_repeated_edges()
    {
        // Same from/to across two call sites should count once.
        var edges = new[] { Edge("A.f", "B.g"), Edge("A.f", "B.g") };
        var result = CouplingComputer.Compute(Modules("A", "B"), edges);
        var a = result.Coupling.Single(c => c.Module == "A");
        Assert.Equal(1, a.Ce);
    }

    [Fact]
    public void ModuleCoupling_instability_balanced_module()
    {
        // C is called once and calls once → I = 1/(1+1) = 0.5.
        var edges = new[] { Edge("A.f", "C.x"), Edge("C.x", "B.g") };
        var result = CouplingComputer.Compute(Modules("A", "B", "C"), edges);
        var c = result.Coupling.Single(m => m.Module == "C");
        Assert.Equal(1, c.Ca);
        Assert.Equal(1, c.Ce);
        Assert.Equal(0.5, c.Instability);
    }
}
```

- [ ] **Step 2: Run tests to verify failure**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~CouplingComputerTests --nologo
```

Expected: FAIL — `CouplingComputer` doesn't exist.

- [ ] **Step 3: Implement CouplingComputer**

Create `src/mcpOffice/Services/Excel/Vba/CouplingComputer.cs`:

```csharp
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class CouplingComputer
{
    public sealed record Result(
        IReadOnlyList<ModuleCoupling> Coupling,
        IReadOnlyList<CouplingPair> Pairs);

    public static Result Compute(
        IReadOnlyList<(string Name, string Kind)> modules,
        IReadOnlyList<ExcelVbaCallEdge> callGraph)
    {
        var ca = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var ce = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var internalEdges = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var pairs = new Dictionary<(string From, string To), int>();

        foreach (var m in modules)
        {
            ca[m.Name] = 0;
            ce[m.Name] = 0;
            internalEdges[m.Name] = 0;
        }

        var seen = new HashSet<(string FromMod, string FromProc, string ToMod, string ToProc)>(EdgeKeyComparer.Instance);

        foreach (var e in callGraph)
        {
            if (!e.Resolved) continue;

            var from = SplitFqn(e.From);
            var to = SplitFqn(e.To);
            if (from is null || to is null) continue;

            var key = (from.Value.Module, from.Value.Procedure, to.Value.Module, to.Value.Procedure);
            if (!seen.Add(key)) continue;

            if (string.Equals(from.Value.Module, to.Value.Module, StringComparison.OrdinalIgnoreCase))
            {
                if (internalEdges.ContainsKey(from.Value.Module))
                    internalEdges[from.Value.Module]++;
            }
            else
            {
                if (ce.ContainsKey(from.Value.Module)) ce[from.Value.Module]++;
                if (ca.ContainsKey(to.Value.Module)) ca[to.Value.Module]++;

                var pairKey = (from.Value.Module, to.Value.Module);
                pairs[pairKey] = pairs.GetValueOrDefault(pairKey, 0) + 1;
            }
        }

        var coupling = modules.Select(m =>
        {
            int ca_ = ca[m.Name];
            int ce_ = ce[m.Name];
            double i = (ca_ + ce_) == 0 ? 0.0 : (double)ce_ / (ca_ + ce_);
            return new ModuleCoupling(m.Name, ca_, ce_, i, internalEdges[m.Name]);
        }).ToList();

        // Pairs computed; sort/produce in Task 8.
        return new Result(coupling, Array.Empty<CouplingPair>());
    }

    private static (string Module, string Procedure)? SplitFqn(string fqn)
    {
        int dot = fqn.IndexOf('.');
        if (dot < 0) return null;
        return (fqn[..dot], fqn[(dot + 1)..]);
    }

    private sealed class EdgeKeyComparer : IEqualityComparer<(string FromMod, string FromProc, string ToMod, string ToProc)>
    {
        public static readonly EdgeKeyComparer Instance = new();
        public bool Equals(
            (string FromMod, string FromProc, string ToMod, string ToProc) a,
            (string FromMod, string FromProc, string ToMod, string ToProc) b) =>
            string.Equals(a.FromMod, b.FromMod, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(a.FromProc, b.FromProc, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(a.ToMod, b.ToMod, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(a.ToProc, b.ToProc, StringComparison.OrdinalIgnoreCase);

        public int GetHashCode((string FromMod, string FromProc, string ToMod, string ToProc) k) =>
            HashCode.Combine(
                k.FromMod.ToLowerInvariant(),
                k.FromProc.ToLowerInvariant(),
                k.ToMod.ToLowerInvariant(),
                k.ToProc.ToLowerInvariant());
    }
}
```

- [ ] **Step 4: Run tests to verify pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~CouplingComputerTests --nologo
```

Expected: PASS — 6/6.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/CouplingComputer.cs tests/mcpOffice.Tests/Excel/Vba/CouplingComputerTests.cs
git commit -m "feat: CouplingComputer — moduleCoupling block"
```

---

### Task 8: CouplingComputer — couplingPairs

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/CouplingComputer.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/CouplingComputerTests.cs`

Pairs already accumulated in the dictionary; this task projects them into `CouplingPair[]`, sorts, and exposes via `Result.Pairs`.

- [ ] **Step 1: Append failing tests**

```csharp
[Fact]
public void Pairs_emit_directional_edge_counts_sorted_desc()
{
    var edges = new[]
    {
        Edge("A.f", "B.g"),
        Edge("A.f", "B.h"),
        Edge("A.k", "B.g"),
        Edge("B.g", "A.f")
    };
    var result = CouplingComputer.Compute(Modules("A", "B"), edges);

    Assert.Equal(2, result.Pairs.Count);
    var ab = result.Pairs.First();
    Assert.Equal("A", ab.From);
    Assert.Equal("B", ab.To);
    Assert.Equal(3, ab.EdgeCount);

    var ba = result.Pairs.Last();
    Assert.Equal("B", ba.From);
    Assert.Equal("A", ba.To);
    Assert.Equal(1, ba.EdgeCount);
}

[Fact]
public void Pairs_omit_zero_count_pairs()
{
    var result = CouplingComputer.Compute(Modules("A", "B", "C"),
        new[] { Edge("A.f", "B.g") });
    // Only A→B is non-zero.
    Assert.Single(result.Pairs);
    Assert.Equal("A", result.Pairs[0].From);
    Assert.Equal("B", result.Pairs[0].To);
}

[Fact]
public void Pairs_stable_sort_alphabetical_within_same_count()
{
    var edges = new[]
    {
        Edge("Z.f", "A.g"),
        Edge("M.f", "B.g")
    };
    var result = CouplingComputer.Compute(Modules("A", "B", "M", "Z"), edges);
    Assert.Equal(2, result.Pairs.Count);
    // Both have edgeCount=1; alphabetical by From: M before Z.
    Assert.Equal("M", result.Pairs[0].From);
    Assert.Equal("Z", result.Pairs[1].From);
}
```

- [ ] **Step 2: Run tests to verify failure**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~CouplingComputerTests --nologo
```

Expected: 3 new tests fail (Pairs is empty).

- [ ] **Step 3: Project and sort the pairs**

Replace `return new Result(coupling, Array.Empty<CouplingPair>());` with:

```csharp
        var pairList = pairs
            .Select(kv => new CouplingPair(kv.Key.From, kv.Key.To, kv.Value))
            .OrderByDescending(p => p.EdgeCount)
            .ThenBy(p => p.From, StringComparer.Ordinal)
            .ThenBy(p => p.To, StringComparer.Ordinal)
            .ToList();

        return new Result(coupling, pairList);
```

- [ ] **Step 4: Run tests to verify pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~CouplingComputerTests --nologo
```

Expected: PASS — 9/9.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/CouplingComputer.cs tests/mcpOffice.Tests/Excel/Vba/CouplingComputerTests.cs
git commit -m "feat: CouplingComputer — couplingPairs"
```

---

# Phase 4 — ParadigmOverlayApplier

### Task 9: Naming + paradigm dispatcher skeleton

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs`

Common naming applies regardless of paradigm: PascalCase + `mod`/`cls`/`frm` prefix strip for class names; PascalCase for method names. The dispatcher routes to per-paradigm rule tables; this task lands the skeleton + naming + an `Apply` for a single trivial classLibrary case so the file is testable.

- [ ] **Step 1: Write the failing tests**

Create `tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs`:

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class ParadigmOverlayApplierTests
{
    private static ProcedureAxes Axes(
        string trigger = "calledOnly",
        string purity = "pure",
        string? shape = "leaf",
        params string[] dependencies) =>
        new(trigger, purity, shape, dependencies);

    [Fact]
    public void Naming_strips_mod_prefix()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "modOrders", procedureName: "ProcessOrder",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("Orders", s.SuggestedClassName);
    }

    [Fact]
    public void Naming_strips_cls_prefix()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "clsCustomer", procedureName: "GetById",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("Customer", s.SuggestedClassName);
    }

    [Fact]
    public void Naming_strips_frm_prefix()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "frmLogin", procedureName: "Validate",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("Login", s.SuggestedClassName);
    }

    [Fact]
    public void Naming_passes_through_when_no_prefix()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "Module1", procedureName: "Main",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("Module1", s.SuggestedClassName);
    }

    [Fact]
    public void Naming_method_pascal_case()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "Module1", procedureName: "do_thing",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("DoThing", s.SuggestedMethodName);
    }

    [Fact]
    public void IsPublic_mirrors_scope()
    {
        var pub = ParadigmOverlayApplier.Apply(
            "M", "P", scope: "Public", axes: Axes(), paradigm: "classLibrary");
        var priv = ParadigmOverlayApplier.Apply(
            "M", "P", scope: "Private", axes: Axes(), paradigm: "classLibrary");
        Assert.True(pub.IsPublic);
        Assert.False(priv.IsPublic);
    }
}
```

- [ ] **Step 2: Run tests to verify failure**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~ParadigmOverlayApplierTests --nologo
```

Expected: FAIL — type doesn't exist.

- [ ] **Step 3: Implement the skeleton**

Create `src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs`:

```csharp
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class ParadigmOverlayApplier
{
    public static readonly IReadOnlyList<string> SupportedParadigms = new[]
    {
        "classLibrary", "workerService", "webApi", "console"
    };

    public static CSharpSuggestion Apply(
        string module,
        string procedureName,
        string? scope,
        ProcedureAxes axes,
        string paradigm)
    {
        var className = StripModulePrefix(module);
        var methodName = ToPascalCase(procedureName);
        var isPublic = !string.Equals(scope, "Private", StringComparison.OrdinalIgnoreCase);

        return paradigm switch
        {
            "classLibrary" => ApplyClassLibrary(className, methodName, isPublic, axes),
            "workerService" => ApplyWorkerService(className, methodName, isPublic, axes),
            "webApi" => ApplyWebApi(className, methodName, isPublic, axes),
            "console" => ApplyConsole(className, methodName, isPublic, axes),
            _ => throw new ArgumentException($"Unsupported paradigm '{paradigm}'", nameof(paradigm))
        };
    }

    // Implemented in Task 10.
    private static CSharpSuggestion ApplyClassLibrary(string c, string m, bool pub, ProcedureAxes axes) =>
        new("staticMethod", c, m, "static", pub, Array.Empty<string>());

    // Implemented in Task 11.
    private static CSharpSuggestion ApplyWorkerService(string c, string m, bool pub, ProcedureAxes axes) =>
        new("instanceMethod", c, m, "scoped", pub, Array.Empty<string>());

    // Implemented in Task 12.
    private static CSharpSuggestion ApplyWebApi(string c, string m, bool pub, ProcedureAxes axes) =>
        new("instanceMethod", c, m, "scoped", pub, Array.Empty<string>());

    // Implemented in Task 13.
    private static CSharpSuggestion ApplyConsole(string c, string m, bool pub, ProcedureAxes axes) =>
        new("staticMethod", c, m, "static", pub, Array.Empty<string>());

    private static string StripModulePrefix(string moduleName)
    {
        foreach (var prefix in new[] { "mod", "cls", "frm" })
        {
            if (moduleName.Length > prefix.Length &&
                moduleName.StartsWith(prefix, StringComparison.Ordinal) &&
                char.IsUpper(moduleName[prefix.Length]))
            {
                return moduleName[prefix.Length..];
            }
        }
        return ToPascalCase(moduleName);
    }

    private static string ToPascalCase(string identifier)
    {
        if (string.IsNullOrEmpty(identifier)) return identifier;
        var parts = identifier.Split('_', StringSplitOptions.RemoveEmptyEntries);
        return string.Concat(parts.Select(p =>
            char.ToUpperInvariant(p[0]) + (p.Length > 1 ? p[1..] : "")));
    }
}
```

- [ ] **Step 4: Run tests to verify pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~ParadigmOverlayApplierTests --nologo
```

Expected: PASS — 6/6.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs
git commit -m "feat: ParadigmOverlayApplier — naming + dispatcher"
```

---

### Task 10: ParadigmOverlayApplier — classLibrary rules

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs`

Implement the design doc's classLibrary table rows:

| Axes | targetType | lifetime | blockers |
|---|---|---|---|
| pure + leaf | staticMethod | static | — |
| pure or readsState, no deps | staticMethod | static | — |
| sideEffectful + (database or network) | instanceMethod | scoped | requires_external_dependency_injection |
| writesState + only excelObjectModel | instanceMethod | scoped | depends_on_excel_object_model |
| eventHandler | requiresManualReview | null | event_handler_no_pure_classlib_target |

- [ ] **Step 1: Append failing tests**

```csharp
[Fact]
public void ClassLib_pure_leaf_static_method_no_blockers()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "P", "Public",
        Axes(purity: "pure", shape: "leaf"), "classLibrary");
    Assert.Equal("staticMethod", s.TargetType);
    Assert.Equal("static", s.Lifetime);
    Assert.Empty(s.Blockers);
}

[Fact]
public void ClassLib_readsState_no_deps_static_method()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "P", "Public",
        Axes(purity: "readsState", shape: "leaf"), "classLibrary");
    Assert.Equal("staticMethod", s.TargetType);
}

[Fact]
public void ClassLib_sideEffectful_database_instance_method_with_blocker()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "P", "Public",
        Axes(purity: "sideEffectful", shape: null, dependencies: new[] { "database" }),
        "classLibrary");
    Assert.Equal("instanceMethod", s.TargetType);
    Assert.Equal("scoped", s.Lifetime);
    Assert.Contains("requires_external_dependency_injection", s.Blockers);
}

[Fact]
public void ClassLib_eventHandler_requires_manual_review()
{
    var s = ParadigmOverlayApplier.Apply("Sheet1", "Worksheet_Change", "Public",
        Axes(trigger: "eventHandler"), "classLibrary");
    Assert.Equal("requiresManualReview", s.TargetType);
    Assert.Null(s.Lifetime);
    Assert.Contains("event_handler_no_pure_classlib_target", s.Blockers);
}
```

- [ ] **Step 2: Run tests — fail**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~ParadigmOverlayApplierTests --nologo
```

Expected: 4 new fail.

- [ ] **Step 3: Implement `ApplyClassLibrary`**

Replace the `ApplyClassLibrary` placeholder:

```csharp
    private static CSharpSuggestion ApplyClassLibrary(string c, string m, bool pub, ProcedureAxes axes)
    {
        if (axes.Trigger == "eventHandler")
            return new("requiresManualReview", c, m, null, pub,
                new[] { "event_handler_no_pure_classlib_target" });

        bool depsEmpty = axes.Dependencies.Count == 0;
        bool excelOnly = axes.Dependencies.Count == 1 && axes.Dependencies[0] == "excelObjectModel";
        bool hasDbOrNet = axes.Dependencies.Any(d => d == "database" || d == "network");

        if (axes.Purity == "pure" && axes.Shape == "leaf")
            return new("staticMethod", c, m, "static", pub, Array.Empty<string>());

        if ((axes.Purity == "pure" || axes.Purity == "readsState") && depsEmpty)
            return new("staticMethod", c, m, "static", pub, Array.Empty<string>());

        if (axes.Purity == "sideEffectful" && hasDbOrNet)
            return new("instanceMethod", c, m, "scoped", pub,
                new[] { "requires_external_dependency_injection" });

        if (axes.Purity == "writesState" && excelOnly)
            return new("instanceMethod", c, m, "scoped", pub,
                new[] { "depends_on_excel_object_model" });

        // Catch-all: instance method, conservative.
        return new("instanceMethod", c, m, "scoped", pub, Array.Empty<string>());
    }
```

- [ ] **Step 4: Run tests — pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~ParadigmOverlayApplierTests --nologo
```

Expected: PASS — 10/10.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs
git commit -m "feat: ParadigmOverlayApplier — classLibrary rules"
```

---

### Task 11: ParadigmOverlayApplier — workerService rules

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs`

Rules:
- macroEntryPoint + (writesState | sideEffectful) → backgroundService, singleton, no blockers (except excelObjectModel append).
- eventHandler whose name matches `Workbook_Open` / `Auto_Open` / `OnTime` → backgroundService, singleton.
- Other → instanceMethod, scoped.
- If `dependencies` includes `excelObjectModel`, append `depends_on_excel_object_model` to blockers.

- [ ] **Step 1: Append failing tests**

```csharp
[Fact]
public void Worker_macroEntryPoint_writesState_becomes_backgroundService()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "Main", "Public",
        Axes(trigger: "macroEntryPoint", purity: "writesState"), "workerService");
    Assert.Equal("backgroundService", s.TargetType);
    Assert.Equal("singleton", s.Lifetime);
}

[Fact]
public void Worker_eventHandler_workbook_open_becomes_backgroundService()
{
    var s = ParadigmOverlayApplier.Apply("ThisWorkbook", "Workbook_Open", "Public",
        Axes(trigger: "eventHandler"), "workerService");
    Assert.Equal("backgroundService", s.TargetType);
}

[Fact]
public void Worker_eventHandler_auto_open_becomes_backgroundService()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "Auto_Open", "Public",
        Axes(trigger: "eventHandler"), "workerService");
    Assert.Equal("backgroundService", s.TargetType);
}

[Fact]
public void Worker_other_procedure_becomes_instanceMethod()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "Helper", "Public",
        Axes(trigger: "calledOnly", purity: "pure"), "workerService");
    Assert.Equal("instanceMethod", s.TargetType);
    Assert.Equal("scoped", s.Lifetime);
}

[Fact]
public void Worker_appends_excel_object_model_blocker()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "Main", "Public",
        Axes(trigger: "macroEntryPoint", purity: "sideEffectful",
             dependencies: new[] { "excelObjectModel" }),
        "workerService");
    Assert.Contains("depends_on_excel_object_model", s.Blockers);
}
```

- [ ] **Step 2: Run tests — fail**

Expected: 5 new fail.

- [ ] **Step 3: Implement `ApplyWorkerService`**

Replace the `ApplyWorkerService` placeholder:

```csharp
    private static CSharpSuggestion ApplyWorkerService(string c, string m, bool pub, ProcedureAxes axes)
    {
        bool isBackgroundEntry =
            (axes.Trigger == "macroEntryPoint" &&
             (axes.Purity == "writesState" || axes.Purity == "sideEffectful")) ||
            (axes.Trigger == "eventHandler" && IsBackgroundEventName(m));

        var blockers = new List<string>();
        if (axes.Dependencies.Contains("excelObjectModel"))
            blockers.Add("depends_on_excel_object_model");

        if (isBackgroundEntry)
            return new("backgroundService", c, m, "singleton", pub, blockers);

        return new("instanceMethod", c, m, "scoped", pub, blockers);
    }

    private static bool IsBackgroundEventName(string method) =>
        string.Equals(method, "WorkbookOpen", StringComparison.Ordinal) ||
        string.Equals(method, "AutoOpen", StringComparison.Ordinal) ||
        method.Contains("OnTime", StringComparison.Ordinal);
```

> Note: `m` is the PascalCased method name (`Workbook_Open` becomes `WorkbookOpen`, `Auto_Open` becomes `AutoOpen`).

- [ ] **Step 4: Run tests — pass**

Expected: PASS — 15/15.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs
git commit -m "feat: ParadigmOverlayApplier — workerService rules"
```

---

### Task 12: ParadigmOverlayApplier — webApi rules

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs`

Rules:
- macroEntryPoint AND isPublic → apiAction, scoped.
- Other → instanceMethod, scoped.
- excelObjectModel dep → append `depends_on_excel_object_model` blocker.

- [ ] **Step 1: Append failing tests**

```csharp
[Fact]
public void WebApi_macroEntryPoint_public_becomes_apiAction()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "GetOrders", "Public",
        Axes(trigger: "macroEntryPoint"), "webApi");
    Assert.Equal("apiAction", s.TargetType);
    Assert.Equal("scoped", s.Lifetime);
}

[Fact]
public void WebApi_macroEntryPoint_private_becomes_instanceMethod()
{
    // Should never happen (private cannot be macroEntryPoint per the trigger rule),
    // but the table still says "any other" for non-public-entry rows.
    var s = ParadigmOverlayApplier.Apply("Module1", "GetOrders", "Private",
        Axes(trigger: "calledOnly"), "webApi");
    Assert.Equal("instanceMethod", s.TargetType);
}

[Fact]
public void WebApi_appends_excel_blocker()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "GetOrders", "Public",
        Axes(trigger: "macroEntryPoint", dependencies: new[] { "excelObjectModel" }),
        "webApi");
    Assert.Contains("depends_on_excel_object_model", s.Blockers);
}
```

- [ ] **Step 2: Run tests — fail**

Expected: 3 new fail.

- [ ] **Step 3: Implement `ApplyWebApi`**

Replace the `ApplyWebApi` placeholder:

```csharp
    private static CSharpSuggestion ApplyWebApi(string c, string m, bool pub, ProcedureAxes axes)
    {
        var blockers = new List<string>();
        if (axes.Dependencies.Contains("excelObjectModel"))
            blockers.Add("depends_on_excel_object_model");

        if (axes.Trigger == "macroEntryPoint" && pub)
            return new("apiAction", c, m, "scoped", pub, blockers);

        return new("instanceMethod", c, m, "scoped", pub, blockers);
    }
```

- [ ] **Step 4: Run tests — pass**

Expected: PASS — 18/18.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs
git commit -m "feat: ParadigmOverlayApplier — webApi rules"
```

---

### Task 13: ParadigmOverlayApplier — console rules

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs`
- Modify: `tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs`

Rules:
- macroEntryPoint → consoleEntryPoint, lifetime null.
- Other: staticMethod (static) if purity ∈ {pure, readsState}; else instanceMethod (scoped).

- [ ] **Step 1: Append failing tests**

```csharp
[Fact]
public void Console_macroEntryPoint_becomes_consoleEntryPoint()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "Main", "Public",
        Axes(trigger: "macroEntryPoint"), "console");
    Assert.Equal("consoleEntryPoint", s.TargetType);
    Assert.Null(s.Lifetime);
}

[Fact]
public void Console_pure_helper_becomes_staticMethod()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "Helper", "Public",
        Axes(trigger: "calledOnly", purity: "pure"), "console");
    Assert.Equal("staticMethod", s.TargetType);
    Assert.Equal("static", s.Lifetime);
}

[Fact]
public void Console_sideEffectful_helper_becomes_instanceMethod()
{
    var s = ParadigmOverlayApplier.Apply("Module1", "Doer", "Public",
        Axes(trigger: "calledOnly", purity: "sideEffectful",
             dependencies: new[] { "filesystem" }),
        "console");
    Assert.Equal("instanceMethod", s.TargetType);
    Assert.Equal("scoped", s.Lifetime);
}
```

- [ ] **Step 2: Run tests — fail**

Expected: 3 new fail.

- [ ] **Step 3: Implement `ApplyConsole`**

Replace the `ApplyConsole` placeholder:

```csharp
    private static CSharpSuggestion ApplyConsole(string c, string m, bool pub, ProcedureAxes axes)
    {
        if (axes.Trigger == "macroEntryPoint")
            return new("consoleEntryPoint", c, m, null, pub, Array.Empty<string>());

        bool isStaticCandidate = axes.Purity == "pure" || axes.Purity == "readsState";
        return isStaticCandidate
            ? new("staticMethod", c, m, "static", pub, Array.Empty<string>())
            : new("instanceMethod", c, m, "scoped", pub, Array.Empty<string>());
    }
```

- [ ] **Step 4: Run tests — pass**

Expected: PASS — 21/21.

- [ ] **Step 5: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/ParadigmOverlayApplier.cs tests/mcpOffice.Tests/Excel/Vba/ParadigmOverlayApplierTests.cs
git commit -m "feat: ParadigmOverlayApplier — console rules"
```

---

# Phase 5 — VbaConversionHintBuilder (orchestrator)

### Task 14: Builder — happy-path orchestration

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaConversionHintBuilder.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaConversionHintBuilderTests.cs`

Wires `AxisClassifier`, `CouplingComputer`, and `ParadigmOverlayApplier`. Validates `targetParadigm`. Filters procedures by `moduleName`. Coupling is always whole-workbook. Generates the rationale string per procedure.

- [ ] **Step 1: Write the failing tests**

Create `tests/mcpOffice.Tests/Excel/Vba/VbaConversionHintBuilderTests.cs`:

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class VbaConversionHintBuilderTests
{
    private static ExcelVbaProcedure Proc(
        string module, string name,
        bool isEventHandler = false, string? scope = "Public") =>
        new(name, $"{module}.{name}", "Sub", scope,
            Array.Empty<ExcelVbaParameter>(), null, 1, 2, isEventHandler, null);

    private static ExcelVbaModuleAnalysis Mod(string name, string kind, params ExcelVbaProcedure[] procs) =>
        new(name, kind, true, null, procs);

    private static ExcelVbaCallEdge Edge(string from, string to, bool resolved = true) =>
        new(from, to, resolved,
            new ExcelVbaSiteRef(from.Split('.')[0], from.Split('.')[1], 1));

    private static ExcelVbaAnalysis MakeAnalysis(
        IReadOnlyList<ExcelVbaModuleAnalysis> modules,
        IReadOnlyList<ExcelVbaCallEdge> callGraph,
        IReadOnlyList<ExcelVbaObjectModelRef>? om = null,
        IReadOnlyList<ExcelVbaDependency>? deps = null)
    {
        var refs = new ExcelVbaReferences(om ?? Array.Empty<ExcelVbaObjectModelRef>(),
                                          deps ?? Array.Empty<ExcelVbaDependency>());
        var procCount = modules.Sum(m => m.Procedures.Count);
        var handlerCount = modules.Sum(m => m.Procedures.Count(p => p.IsEventHandler));
        var summary = new ExcelVbaAnalysisSummary(
            modules.Count, modules.Count, 0, procCount, handlerCount,
            callGraph.Count, refs.ObjectModel.Count, refs.Dependencies.Count);
        return new ExcelVbaAnalysis(true, summary, modules, callGraph, refs);
    }

    [Fact]
    public void Builder_emits_one_hint_per_procedure()
    {
        var analysis = MakeAnalysis(
            new[]
            {
                Mod("Module1", "standard", Proc("Module1", "A"), Proc("Module1", "B")),
                Mod("Module2", "standard", Proc("Module2", "C"))
            },
            Array.Empty<ExcelVbaCallEdge>());

        var result = VbaConversionHintBuilder.Build(analysis, moduleName: null, targetParadigm: null);
        Assert.Equal(3, result.ProcedureHints.Count);
        Assert.Equal(3, result.Summary.TotalProcedures);
        Assert.Equal(3, result.Summary.HintedProcedures);
        Assert.Equal(2, result.Summary.ModuleCount);
    }

    [Fact]
    public void Builder_filters_hints_by_moduleName_case_insensitive()
    {
        var analysis = MakeAnalysis(
            new[]
            {
                Mod("Module1", "standard", Proc("Module1", "A")),
                Mod("Module2", "standard", Proc("Module2", "C"))
            },
            Array.Empty<ExcelVbaCallEdge>());

        var result = VbaConversionHintBuilder.Build(analysis, moduleName: "MODULE1", targetParadigm: null);
        Assert.Single(result.ProcedureHints);
        Assert.Equal("Module1", result.ProcedureHints[0].Module);
        Assert.Equal(1, result.Summary.HintedProcedures);
        Assert.Equal(2, result.Summary.TotalProcedures);
    }

    [Fact]
    public void Builder_throws_module_not_found_for_unknown_filter()
    {
        var analysis = MakeAnalysis(
            new[] { Mod("Module1", "standard", Proc("Module1", "A")) },
            Array.Empty<ExcelVbaCallEdge>());
        var act = () => VbaConversionHintBuilder.Build(analysis, moduleName: "Nope", targetParadigm: null);
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("[module_not_found]", ex.Message);
    }

    [Fact]
    public void Builder_throws_unsupported_paradigm()
    {
        var analysis = MakeAnalysis(
            new[] { Mod("Module1", "standard", Proc("Module1", "A")) },
            Array.Empty<ExcelVbaCallEdge>());
        var act = () => VbaConversionHintBuilder.Build(analysis, moduleName: null, targetParadigm: "blazor");
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("[unsupported_paradigm]", ex.Message);
        Assert.Contains("blazor", ex.Message);
    }

    [Fact]
    public void Builder_emits_csharpSuggestion_only_when_paradigm_provided()
    {
        var analysis = MakeAnalysis(
            new[] { Mod("Module1", "standard", Proc("Module1", "A")) },
            Array.Empty<ExcelVbaCallEdge>());

        var without = VbaConversionHintBuilder.Build(analysis, null, null);
        Assert.Null(without.ProcedureHints[0].CsharpSuggestion);

        var with = VbaConversionHintBuilder.Build(analysis, null, "classLibrary");
        Assert.NotNull(with.ProcedureHints[0].CsharpSuggestion);
    }

    [Fact]
    public void Builder_coupling_always_whole_workbook_even_with_module_filter()
    {
        var analysis = MakeAnalysis(
            new[]
            {
                Mod("A", "standard", Proc("A", "f")),
                Mod("B", "standard", Proc("B", "g")),
                Mod("C", "standard", Proc("C", "h"))
            },
            new[] { Edge("A.f", "B.g"), Edge("B.g", "C.h") });

        var result = VbaConversionHintBuilder.Build(analysis, moduleName: "A", targetParadigm: null);
        Assert.Equal(3, result.ModuleCoupling.Count);  // all modules present
        Assert.Equal(2, result.CouplingPairs.Count);   // A→B, B→C
    }

    [Fact]
    public void Builder_caller_callee_counts_match_call_graph()
    {
        var analysis = MakeAnalysis(
            new[]
            {
                Mod("M", "standard",
                    Proc("M", "Caller", scope: "Public"),
                    Proc("M", "Callee", scope: "Public"))
            },
            new[] { Edge("M.Caller", "M.Callee") });

        var result = VbaConversionHintBuilder.Build(analysis, null, null);
        var caller = result.ProcedureHints.Single(p => p.ProcedureName == "Caller");
        var callee = result.ProcedureHints.Single(p => p.ProcedureName == "Callee");
        Assert.Equal(0, caller.CallerCount);
        Assert.Equal(1, caller.CalleeCount);
        Assert.Equal(1, callee.CallerCount);
        Assert.Equal(0, callee.CalleeCount);
    }

    [Fact]
    public void Builder_returns_empty_hints_when_no_vba_project()
    {
        var empty = new ExcelVbaAnalysis(
            HasVbaProject: false,
            Summary: new ExcelVbaAnalysisSummary(0, 0, 0, 0, 0, 0, 0, 0),
            Modules: null, CallGraph: null, References: null);
        var result = VbaConversionHintBuilder.Build(empty, null, null);
        Assert.Empty(result.ProcedureHints);
        Assert.Empty(result.ModuleCoupling);
        Assert.Empty(result.CouplingPairs);
    }
}
```

- [ ] **Step 2: Run tests — fail**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~VbaConversionHintBuilderTests --nologo
```

Expected: FAIL — type doesn't exist.

- [ ] **Step 3: Implement the builder**

Create `src/mcpOffice/Services/Excel/Vba/VbaConversionHintBuilder.cs`:

```csharp
using System.Diagnostics;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class VbaConversionHintBuilder
{
    public static ConversionHints Build(
        ExcelVbaAnalysis analysis,
        string? moduleName,
        string? targetParadigm)
    {
        var sw = Stopwatch.StartNew();

        if (targetParadigm is not null &&
            !ParadigmOverlayApplier.SupportedParadigms.Contains(targetParadigm))
        {
            throw ToolError.UnsupportedParadigm(targetParadigm, ParadigmOverlayApplier.SupportedParadigms);
        }

        if (!analysis.HasVbaProject || analysis.Modules is null)
        {
            return new ConversionHints(
                Summary: new ConversionHintsSummary(0, 0, 0, targetParadigm, sw.ElapsedMilliseconds),
                ProcedureHints: Array.Empty<ProcedureHint>(),
                ModuleCoupling: Array.Empty<ModuleCoupling>(),
                CouplingPairs: Array.Empty<CouplingPair>());
        }

        var modules = analysis.Modules!;
        var callGraph = analysis.CallGraph ?? Array.Empty<ExcelVbaCallEdge>();
        var objectModel = analysis.References?.ObjectModel ?? Array.Empty<ExcelVbaObjectModelRef>();
        var dependencies = analysis.References?.Dependencies ?? Array.Empty<ExcelVbaDependency>();

        // Resolve moduleName (case-insensitive) to canonical name; throw if unknown.
        string? canonicalFilter = null;
        if (!string.IsNullOrWhiteSpace(moduleName))
        {
            var match = modules.FirstOrDefault(m =>
                string.Equals(m.Name, moduleName, StringComparison.OrdinalIgnoreCase));
            if (match is null)
            {
                throw ToolError.ModuleNotFound(moduleName, modules.Select(m => m.Name));
            }
            canonicalFilter = match.Name;
        }

        // Coupling — always whole-workbook.
        var moduleNamesAndKinds = modules.Select(m => (m.Name, m.Kind)).ToList();
        var couplingResult = CouplingComputer.Compute(moduleNamesAndKinds, callGraph);

        // Hints — filtered by moduleName when set.
        var hints = new List<ProcedureHint>();
        int totalProcedures = 0;

        foreach (var mod in modules)
        {
            totalProcedures += mod.Procedures.Count;

            if (canonicalFilter is not null && !string.Equals(mod.Name, canonicalFilter, StringComparison.Ordinal))
                continue;

            foreach (var proc in mod.Procedures)
            {
                var axes = AxisClassifier.Classify(proc, mod.Kind, callGraph, objectModel, dependencies);
                int callerCount = callGraph.Count(e =>
                    string.Equals(e.To, proc.FullyQualifiedName, StringComparison.OrdinalIgnoreCase));
                int calleeCount = callGraph.Count(e =>
                    string.Equals(e.From, proc.FullyQualifiedName, StringComparison.OrdinalIgnoreCase));

                CSharpSuggestion? suggestion = null;
                if (targetParadigm is not null)
                {
                    suggestion = ParadigmOverlayApplier.Apply(
                        mod.Name, proc.Name, proc.Scope, axes, targetParadigm);
                }

                var rationale = BuildRationale(proc, axes, suggestion);

                hints.Add(new ProcedureHint(
                    Module: mod.Name,
                    ProcedureName: proc.Name,
                    Kind: proc.Kind,
                    IsEventHandler: proc.IsEventHandler,
                    ParamCount: proc.Parameters.Count,
                    CallerCount: callerCount,
                    CalleeCount: calleeCount,
                    Axes: axes,
                    Rationale: rationale,
                    CsharpSuggestion: suggestion));
            }
        }

        var summary = new ConversionHintsSummary(
            TotalProcedures: totalProcedures,
            HintedProcedures: hints.Count,
            ModuleCount: modules.Count,
            TargetParadigm: targetParadigm,
            WallTimeMs: sw.ElapsedMilliseconds);

        return new ConversionHints(summary, hints, couplingResult.Coupling, couplingResult.Pairs);
    }

    private static string BuildRationale(
        ExcelVbaProcedure proc, ProcedureAxes axes, CSharpSuggestion? suggestion)
    {
        var parts = new List<string>
        {
            $"{axes.Purity} {(axes.Shape ?? "")}".Trim(),
            $"trigger={axes.Trigger}",
            $"params={proc.Parameters.Count}"
        };
        if (axes.Dependencies.Count > 0)
            parts.Add($"deps=[{string.Join(",", axes.Dependencies)}]");

        var baseLine = string.Join("; ", parts);

        if (suggestion is null) return baseLine;

        var paradigmHint = suggestion.TargetType == "requiresManualReview"
            ? $"Manual review required: {string.Join(", ", suggestion.Blockers)}."
            : $"Suggested as {suggestion.TargetType} on {suggestion.SuggestedClassName}.{suggestion.SuggestedMethodName}.";

        return $"{baseLine} — {paradigmHint}";
    }
}
```

- [ ] **Step 4: Run tests — pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~VbaConversionHintBuilderTests --nologo
```

Expected: PASS — 8/8.

- [ ] **Step 5: Run the whole test suite**

```
dotnet test --nologo
```

Expected: PASS overall.

- [ ] **Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaConversionHintBuilder.cs tests/mcpOffice.Tests/Excel/Vba/VbaConversionHintBuilderTests.cs
git commit -m "feat: VbaConversionHintBuilder — orchestrator"
```

---

# Phase 6 — Service & Tool surface

### Task 15: Add `SuggestVbaConversion` to the service

**Files:**
- Modify: `src/mcpOffice/Services/Excel/IExcelWorkbookService.cs`
- Modify: `src/mcpOffice/Services/Excel/ExcelWorkbookService.cs`

The service method runs the analyzer **without** a moduleName filter (so the coupling block sees the whole graph) and hands off to the builder, which applies the filter.

- [ ] **Step 1: Extend the interface**

Add to `IExcelWorkbookService.cs`:

```csharp
    ConversionHints SuggestVbaConversion(
        string path,
        string? moduleName,
        string? targetParadigm);
```

- [ ] **Step 2: Run build — fails (interface not implemented)**

```
dotnet build --nologo
```

Expected: FAIL — `'ExcelWorkbookService' does not implement interface member`.

- [ ] **Step 3: Locate where `AnalyzeVba` is implemented in `ExcelWorkbookService.cs`**

```
grep -n "AnalyzeVba\|RenderVbaCallgraph\|VbaSourceAnalyzer.Analyze" src/mcpOffice/Services/Excel/ExcelWorkbookService.cs
```

You'll see `AnalyzeVba` calling `VbaProjectReader.Read` then `VbaSourceAnalyzer.Analyze`. Mirror that shape.

- [ ] **Step 4: Implement `SuggestVbaConversion`**

Add to `ExcelWorkbookService.cs` (alongside `AnalyzeVba` and `RenderVbaCallgraph`):

```csharp
    public ConversionHints SuggestVbaConversion(
        string path,
        string? moduleName,
        string? targetParadigm)
    {
        PathGuard.RequireExists(path);
        try
        {
            // Read project + run full analyzer (no module filter — coupling needs whole graph).
            var project = VbaProjectReader.Read(path);
            var analysis = VbaSourceAnalyzer.Analyze(
                project,
                includeProcedures: true,
                includeCallGraph: true,
                includeReferences: true,
                moduleName: null);

            return VbaConversionHintBuilder.Build(analysis, moduleName, targetParadigm);
        }
        catch (Exception ex) when (ex is not ModelContextProtocol.McpException)
        {
            throw ToolError.VbaParseError(path, ex.Message);
        }
    }
```

> **Note:** if `VbaProjectReader.Read` is named differently or takes more parameters in this codebase, match the exact signature from the existing `AnalyzeVba` method.

- [ ] **Step 5: Run build — succeeds**

```
dotnet build --nologo
```

Expected: 0 warnings, 0 errors.

- [ ] **Step 6: Run all tests — succeeds**

```
dotnet test --nologo
```

Expected: PASS overall (no new tests yet — service untested directly until the synthetic and Air tests in later tasks).

- [ ] **Step 7: Commit**

```bash
git add src/mcpOffice/Services/Excel/IExcelWorkbookService.cs src/mcpOffice/Services/Excel/ExcelWorkbookService.cs
git commit -m "feat: IExcelWorkbookService.SuggestVbaConversion"
```

---

### Task 16: Expose `excel_suggest_vba_conversion` MCP tool

**Files:**
- Modify: `src/mcpOffice/Tools/ExcelTools.cs`

- [ ] **Step 1: Add the tool method**

Append to `ExcelTools.cs` after `ExcelRenderVbaCallgraph`:

```csharp
    [McpServerTool(Name = "excel_suggest_vba_conversion")]
    [Description("Conversion-hints layer over excel_analyze_vba. For each VBA procedure, emits multi-axis classification (trigger / purity / shape / dependencies), a human-readable rationale, and — when targetParadigm is set — a structured C# emission target (targetType / class / method / lifetime / blockers). Also returns workbook-wide module coupling: per-module Ca/Ce/instability + pairwise edge counts. moduleName scopes hints to a single module; coupling stays whole-workbook regardless. targetParadigm must be one of classLibrary, workerService, webApi, console.")]
    public static object ExcelSuggestVbaConversion(
        [Description("Absolute path to the .xlsm/.xlsb workbook")] string path,
        [Description("Optional case-insensitive VBA module name to scope per-procedure hints to. Coupling stays whole-workbook. Throws module_not_found if unknown.")] string? moduleName = null,
        [Description("Optional target paradigm: classLibrary | workerService | webApi | console. When set, every hint includes a structured csharpSuggestion. Throws unsupported_paradigm if the value is not in the supported set.")] string? targetParadigm = null)
        => Service.SuggestVbaConversion(path, moduleName, targetParadigm);
```

- [ ] **Step 2: Run build**

```
dotnet build --nologo
```

Expected: 0 warnings, 0 errors.

- [ ] **Step 3: Smoke-run the server**

```powershell
echo "" | dotnet run --project src/mcpOffice --no-build
```

Expected: process starts and exits gracefully on EOF (no exceptions).

- [ ] **Step 4: Commit**

```bash
git add src/mcpOffice/Tools/ExcelTools.cs
git commit -m "feat: excel_suggest_vba_conversion MCP tool"
```

---

# Phase 7 — Cross-cutting & integration

### Task 17: Bump tool-surface integration test (25 → 26)

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`

- [ ] **Step 1: Add the new tool name to the expected array**

Insert `"excel_suggest_vba_conversion",` into the alphabetically-sorted list (between `excel_render_vba_callgraph` and `Ping`):

```csharp
            "excel_render_vba_callgraph",
            "excel_suggest_vba_conversion",
            "Ping",
```

- [ ] **Step 2: Run the integration test — pass**

```
dotnet test tests/mcpOffice.Tests.Integration --filter FullyQualifiedName~ToolSurfaceTests --nologo
```

Expected: PASS.

- [ ] **Step 3: Commit**

```bash
git add tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs
git commit -m "test: tool-surface canary covers excel_suggest_vba_conversion"
```

---

### Task 18: Synthetic-fixture end-to-end test

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Vba/SyntheticConversionHintsTests.cs`

Reuses `tests/fixtures/synthetic-vba.xlsm`. Mirrors `SyntheticAnalyzeTests`. Asserts the service's full pipeline produces a populated payload.

- [ ] **Step 1: Write the test**

Create `tests/mcpOffice.Tests/Excel/Vba/SyntheticConversionHintsTests.cs`:

```csharp
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

public class SyntheticConversionHintsTests
{
    [Fact]
    public void Suggests_conversion_for_synthetic_workbook()
    {
        var path = TestFixtures.Path("synthetic-vba.xlsm");
        var svc = new ExcelWorkbookService();

        var result = svc.SuggestVbaConversion(path, moduleName: null, targetParadigm: null);

        Assert.True(result.Summary.TotalProcedures > 0);
        Assert.Equal(result.Summary.TotalProcedures, result.Summary.HintedProcedures);
        Assert.Equal(4, result.Summary.ModuleCount);                          // synthetic-vba has 4 modules
        Assert.NotEmpty(result.ProcedureHints);
        Assert.NotEmpty(result.ModuleCoupling);

        // Every hint has a non-null axes object and rationale.
        Assert.All(result.ProcedureHints, h =>
        {
            Assert.NotNull(h.Axes);
            Assert.False(string.IsNullOrWhiteSpace(h.Rationale));
            Assert.Null(h.CsharpSuggestion);  // no paradigm requested
        });
    }

    [Fact]
    public void Suggests_conversion_with_classLibrary_paradigm_populates_csharpSuggestion()
    {
        var path = TestFixtures.Path("synthetic-vba.xlsm");
        var svc = new ExcelWorkbookService();

        var result = svc.SuggestVbaConversion(path, moduleName: null, targetParadigm: "classLibrary");

        Assert.Equal("classLibrary", result.Summary.TargetParadigm);
        Assert.All(result.ProcedureHints, h => Assert.NotNull(h.CsharpSuggestion));
    }

    [Fact]
    public void Suggests_conversion_filtered_by_module_name()
    {
        var path = TestFixtures.Path("synthetic-vba.xlsm");
        var svc = new ExcelWorkbookService();

        // Discover a module name first via the analyzer.
        var analysis = svc.AnalyzeVba(path, includeProcedures: true, includeCallGraph: false, includeReferences: false);
        var firstModule = analysis.Modules!.First(m => m.Procedures.Count > 0).Name;

        var result = svc.SuggestVbaConversion(path, moduleName: firstModule, targetParadigm: null);

        Assert.True(result.Summary.HintedProcedures < result.Summary.TotalProcedures
                    || result.Summary.ModuleCount == 1);
        Assert.All(result.ProcedureHints, h => Assert.Equal(firstModule, h.Module));
        Assert.True(result.ModuleCoupling.Count >= 1);   // whole-workbook coupling unaffected by filter
    }
}
```

- [ ] **Step 2: Run tests — pass**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~SyntheticConversionHintsTests --nologo
```

Expected: PASS — 3/3.

- [ ] **Step 3: Commit**

```bash
git add tests/mcpOffice.Tests/Excel/Vba/SyntheticConversionHintsTests.cs
git commit -m "test: synthetic conversion-hints end-to-end test"
```

---

### Task 19: Air sample real-world test (gated)

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Vba/AirSampleConversionHintsTests.cs`

- [ ] **Step 1: Write the test**

Create `tests/mcpOffice.Tests/Excel/Vba/AirSampleConversionHintsTests.cs`:

```csharp
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

public class AirSampleConversionHintsTests
{
    private const string SamplePath = @"C:\Projects\mcpOffice-samples\Air.xlsm";

    [Fact]
    public void Suggests_conversion_for_air_workbook_within_budget()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var result = svc.SuggestVbaConversion(SamplePath, moduleName: null, targetParadigm: null);
        sw.Stop();

        // Plausible floors against the documented Air.xlsm metrics
        // (107 modules, 200 procedures).
        Assert.True(result.Summary.ModuleCount > 50,
            $"expected > 50 modules, got {result.Summary.ModuleCount}");
        Assert.True(result.Summary.TotalProcedures > 100,
            $"expected > 100 procedures, got {result.Summary.TotalProcedures}");
        Assert.True(result.Summary.TotalProcedures == result.Summary.HintedProcedures);

        // Every procedure has a hint.
        Assert.Equal(result.Summary.TotalProcedures, result.ProcedureHints.Count);

        // Coupling block populated.
        Assert.True(result.ModuleCoupling.Count >= result.Summary.ModuleCount);
        Assert.NotEmpty(result.CouplingPairs);

        // Performance budget — generous to absorb cold-cache variance.
        Assert.True(sw.ElapsedMilliseconds < 600,
            $"expected < 600 ms, got {sw.ElapsedMilliseconds} ms");
    }

    [Fact]
    public void ClassLibrary_paradigm_produces_static_methods_and_manual_reviews()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var result = svc.SuggestVbaConversion(SamplePath, null, "classLibrary");

        Assert.Contains(result.ProcedureHints,
            h => h.CsharpSuggestion?.TargetType == "staticMethod");
        Assert.Contains(result.ProcedureHints,
            h => h.CsharpSuggestion?.TargetType == "requiresManualReview");
    }
}
```

- [ ] **Step 2: Run tests — pass on machines with the sample, no-op elsewhere**

```
dotnet test tests/mcpOffice.Tests --filter FullyQualifiedName~AirSampleConversionHintsTests --nologo
```

Expected: PASS (locally on dev machine that has Air.xlsm). No-op on machines without it.

- [ ] **Step 3: Commit**

```bash
git add tests/mcpOffice.Tests/Excel/Vba/AirSampleConversionHintsTests.cs
git commit -m "test: gated Air.xlsm conversion-hints benchmark"
```

---

### Task 20: One stdio integration test

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs`

A round-trip happy path through the JSON-RPC layer, mirroring the existing Excel workflow tests.

- [ ] **Step 1: Inspect the existing pattern**

```
head -80 tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs
```

Look at how an existing test like `Analyzes_vba_via_stdio` calls `harness.Client.CallToolAsync`, copies the synthetic fixture, and asserts JSON shape. Mirror that.

- [ ] **Step 2: Add the test**

Append a new fact to `ExcelWorkflowTests.cs`. Names in this codebase use the `_via_stdio` suffix convention. Use the synthetic fixture (already in `tests/fixtures/synthetic-vba.xlsm`) so the test is unconditional.

```csharp
    [Fact]
    public async Task Suggests_vba_conversion_via_stdio()
    {
        var fixture = TestFixtures.Path("synthetic-vba.xlsm");

        await using var harness = await ServerHarness.StartAsync();
        var response = await harness.Client.CallToolAsync(
            "excel_suggest_vba_conversion",
            new Dictionary<string, object?>
            {
                ["path"] = fixture,
                ["targetParadigm"] = "classLibrary"
            });

        var text = response.Content
            .OfType<ModelContextProtocol.Protocol.TextContentBlock>()
            .Single().Text;

        // Loose JSON-shape assertions — the unit tests cover the matrix; this just
        // verifies the JSON-RPC layer doesn't drop fields.
        Assert.Contains("\"summary\"", text);
        Assert.Contains("\"procedureHints\"", text);
        Assert.Contains("\"moduleCoupling\"", text);
        Assert.Contains("\"couplingPairs\"", text);
        Assert.Contains("\"csharpSuggestion\"", text);
        Assert.Contains("\"targetParadigm\":\"classLibrary\"", text);
    }
```

> If the test file uses a `TestFixtures` helper from a different namespace than the unit tests, mirror what other Excel integration tests do (the Word integration tests have a similar fixture path pattern).

- [ ] **Step 3: Run integration tests — pass**

```
dotnet test tests/mcpOffice.Tests.Integration --filter FullyQualifiedName~Suggests_vba_conversion_via_stdio --nologo
```

Expected: PASS.

- [ ] **Step 4: Run the entire integration suite — sanity**

```
dotnet test tests/mcpOffice.Tests.Integration --nologo
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs
git commit -m "test: stdio round-trip for excel_suggest_vba_conversion"
```

---

# Phase 8 — Final verification

### Task 21: Release-mode verification + handoff

**Files:**
- None new; verifies and updates handoff docs.

- [ ] **Step 1: Clean Release build**

```
dotnet build -c Release --nologo
```

Expected: 0 warnings, 0 errors.

- [ ] **Step 2: Release test pass**

```
dotnet test -c Release --nologo
```

Expected: ALL PASS — count should be at least 208 (existing) + ~37 new (axes 19 + coupling 9 + overlay 21 + builder 8 + synthetic 3 + 1 new error test - some adjustments) on the unit side, plus 1 new integration test (now 14 total).

- [ ] **Step 3: Live verify the new tool against `synthetic-vba.xlsm`**

Wire the Release build into a Claude Code session by editing your `.mcp.json` (or local config) to point at `src/mcpOffice/bin/Release/net9.0/mcpOffice.exe`, restart, and call:

```
excel_suggest_vba_conversion(path: "<repo>\\tests\\fixtures\\synthetic-vba.xlsm", targetParadigm: "classLibrary")
```

Expected: returns a JSON object with `summary`, `procedureHints[]` (multiple entries with `csharpSuggestion` populated), `moduleCoupling[]`, `couplingPairs[]`. Skim the output for sanity — module names match, axes look reasonable.

> Per global CLAUDE.md: green build + green tests is not "it works." Do this live verification before considering the task done.

- [ ] **Step 4: Update `SESSION_HANDOFF.md` and `TODO.md`**

In `SESSION_HANDOFF.md`, replace the current "Where Things Stand" + "Next Up" sections with an entry describing v3 landing. Mention:
- Branch state (assuming PR merged): `main` clean, latest commit is the squash-merged v3 PR.
- Tool surface count: 25 → 26.
- New tool: `excel_suggest_vba_conversion`.
- Two follow-ups still on TODO: cluster detection, pagination on `procedureHints[]`.

In `TODO.md`, mark the v3 conversion-hints layer as DONE (mirror the `excel_analyze_vba v1 + v2 — DONE` block style). Move the deferred items (cluster detection, paradigm: blazor/winforms/wpf, cyclomatic complexity, module-scope-write detection) into the "Side items" section under a new heading.

- [ ] **Step 5: Commit handoff updates**

```bash
git add SESSION_HANDOFF.md TODO.md
git commit -m "chore: handoff after analyzer v3"
```

- [ ] **Step 6: Open a PR**

```bash
git push -u origin feat/excel-vba-conversion-hints-v3
gh pr create --title "feat: excel_suggest_vba_conversion (analyzer v3)" --body "$(cat <<'EOF'
## Summary
- Adds `excel_suggest_vba_conversion` (26th MCP tool) — a conversion-hints layer over `excel_analyze_vba`.
- Per-procedure multi-axis classification (trigger / purity / shape / dependencies) plus optional `targetParadigm` overlay producing structured C# emission targets for one of `classLibrary`, `workerService`, `webApi`, `console`.
- Workbook-wide module coupling: per-module `Ca`/`Ce`/`instability`/`internalEdges` and directional `couplingPairs`.
- No new VBA parsing — every hint derived from v1's existing `ExcelVbaAnalysis`.

Design: `docs/plans/2026-05-07-mcpoffice-excel-analyze-vba-v3-design.md`.
Plan: `docs/plans/2026-05-07-mcpoffice-excel-analyze-vba-v3-plan.md`.

## Test plan
- [ ] `dotnet build -c Release` clean, 0 warnings.
- [ ] `dotnet test -c Release` all green.
- [ ] Synthetic-fixture end-to-end test passes (`SyntheticConversionHintsTests`).
- [ ] Air.xlsm gated benchmark passes locally with wall time < 600 ms.
- [ ] Live stdio verification via Claude Code against `tests/fixtures/synthetic-vba.xlsm`.
EOF
)"
```

Expected: PR created. Capture the URL.

- [ ] **Step 7: Final task list update — mark all done**

(Plan complete.)

---

## What this plan deliberately does NOT do

- **No cluster detection.** Pairs are the substrate; clustering layers on top later.
- **No pagination on `procedureHints[]`.** Same problem as the analyzer's heavy arrays — solve once for both.
- **No module-scope-write regex.** The design doc flagged this as an open question; this plan defers — `purity` collapses to 3 of 4 values (`pure` / `readsState` / `sideEffectful`). `writesState` will activate when `ExcelVbaObjectModelRef.Mode` lands.
- **No `blazor` / `winforms` / `wpf` paradigms.** Form-layout analysis is out of scope.
- **No cyclomatic complexity.** Needs a deeper VBA parser.

## Risks called out

1. **`ExcelVbaObjectModelRef.Mode` absent in v1.** The design doc's purity rules assume read/write tagging on object-model refs. Current record doesn't expose `Mode`. The plan downgrades the `writesState` path to never fire — every object-model touch is `readsState`. If this proves too coarse against Air.xlsm, follow-up work to add `Mode` to v1's collector.
2. **`ExcelVbaDependency.Kind` may use values outside the design's closed set.** v1 emits `automation` for shell-out / `Application.Run`; the plan maps it to `shell`. If other unexpected kinds surface (e.g. `oledb`, `wmi`), the dependencies axis will pass them through verbatim — pure noise for the agent, but won't crash. Spike during the synthetic test; tighten the mapping if needed.
3. **`VbaProjectReader.Read` signature.** The plan assumes the same shape `AnalyzeVba` already uses. If it differs, mirror the actual existing `AnalyzeVba` implementation when wiring `SuggestVbaConversion`.
4. **Performance budget.** The plan targets < 200 ms on Air.xlsm but the gated test allows 600 ms to absorb cold-cache and CI-machine variance. If a real measurement comes in over 200 ms, profile `BuildRationale` and the linq-heavy loops in the builder before optimising the analyzer.

## Self-review

**Spec coverage check (against `docs/plans/2026-05-07-mcpoffice-excel-analyze-vba-v3-design.md`):**

- Tool surface (`excel_suggest_vba_conversion`, 3 params): Tasks 15–16. ✓
- Output schema (Summary / ProcedureHints / ModuleCoupling / CouplingPairs): Task 2 + builder Tasks 14, 15. ✓
- Axes — trigger: Task 3. ✓
- Axes — purity: Task 4 (with documented downgrade). ✓
- Axes — shape: Task 5. ✓
- Axes — dependencies: Task 6. ✓
- Coupling — `moduleCoupling`: Task 7. ✓
- Coupling — `couplingPairs`: Task 8. ✓
- Naming convention (mod/cls/frm strip + PascalCase): Task 9. ✓
- Paradigm — classLibrary rules: Task 10. ✓
- Paradigm — workerService rules: Task 11. ✓
- Paradigm — webApi rules: Task 12. ✓
- Paradigm — console rules: Task 13. ✓
- Blocker codes: Tasks 10–12. ✓
- `unsupported_paradigm` error: Task 1, validated in Task 14. ✓
- `module_not_found` reuse: Task 14. ✓
- Performance reporting (`wallTimeMs`): Task 14. ✓
- Synthetic test: Task 18. ✓
- Air sample test: Task 19. ✓
- Tool surface canary: Task 17. ✓
- One stdio integration test: Task 20. ✓
- Final live verification: Task 21. ✓

**Type consistency check:** `ProcedureAxes`, `ProcedureHint`, `CSharpSuggestion`, `ModuleCoupling`, `CouplingPair`, `ConversionHints`, `ConversionHintsSummary` are defined in Task 2 and used identically in Tasks 3–14, 18, 19, 20. `ParadigmOverlayApplier.SupportedParadigms` defined in Task 9, used in Task 14. `CouplingComputer.Compute` returns a nested `Result` record consumed by Task 14.

**Placeholder scan:** No "TODO", "implement later", "fill in details", "appropriate error handling", "similar to Task N" — every code block is concrete. The two `> Note:` callouts in Tasks 15 and 20 ask the implementer to mirror an existing signature; that's a verification step, not a placeholder.
