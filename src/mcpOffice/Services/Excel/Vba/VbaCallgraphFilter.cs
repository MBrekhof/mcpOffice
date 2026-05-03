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
}
