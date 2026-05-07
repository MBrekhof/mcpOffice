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
}
