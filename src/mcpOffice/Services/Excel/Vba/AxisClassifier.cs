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
