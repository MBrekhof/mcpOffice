using McpOffice;
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
    }
}
