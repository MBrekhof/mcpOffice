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
        var purity = ClassifyPurity(proc, objectModel, dependencies);
        var shape = ClassifyShape(proc, callGraph);
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

    private static string ProcModule(ExcelVbaProcedure proc) =>
        proc.FullyQualifiedName.Split('.', 2)[0];
}
