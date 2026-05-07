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
