using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

public sealed record CallgraphFilterOptions(
    string? ModuleName = null,
    string? ProcedureName = null,
    int Depth = 2,
    string Direction = "both",
    int MaxNodes = 300);

public static class VbaCallgraphFilter
{
    private const string ExternalIdPrefix = "__ext__";

    private static string ExternalId(string calleeName) => ExternalIdPrefix + calleeName;

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

        var allNodesById = new Dictionary<string, CallgraphNode>(StringComparer.Ordinal);
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

        var allEdges = analysis.CallGraph ?? (IReadOnlyList<ExcelVbaCallEdge>)Array.Empty<ExcelVbaCallEdge>();

        if (!string.IsNullOrWhiteSpace(options.ProcedureName))
        {
            if (moduleFilter is null)
                throw ToolError.InvalidRenderOption(
                    $"procedureName='{options.ProcedureName}' requires moduleName — bare procedure names aren't unique.");

            var moduleProcs = analysis.Modules.Single(m => m.Name == moduleFilter).Procedures;
            var focalMatch = moduleProcs.FirstOrDefault(p =>
                string.Equals(p.Name, options.ProcedureName, StringComparison.OrdinalIgnoreCase));
            if (focalMatch is null)
                throw ToolError.ProcedureNotFound(options.ProcedureName, moduleProcs.Select(p => p.Name));

            var focalId = focalMatch.FullyQualifiedName;
            var followCallees = options.Direction is "callees" or "both";
            var followCallers = options.Direction is "callers" or "both";
            if (!followCallees && !followCallers)
                throw ToolError.InvalidRenderOption(
                    $"direction='{options.Direction}' is not one of callees, callers, both.");

            var visited = new HashSet<string>(StringComparer.Ordinal) { focalId };
            var frontier = new HashSet<string>(StringComparer.Ordinal) { focalId };
            for (var hop = 0; hop < options.Depth; hop++)
            {
                var next = new HashSet<string>(StringComparer.Ordinal);
                foreach (var e in allEdges)
                {
                    if (followCallees && frontier.Contains(e.From) && !visited.Contains(e.To)
                        && allNodesById.ContainsKey(e.To))
                        next.Add(e.To);
                    if (followCallers && frontier.Contains(e.To) && !visited.Contains(e.From)
                        && allNodesById.ContainsKey(e.From))
                        next.Add(e.From);
                }
                if (next.Count == 0) break;
                foreach (var id in next) visited.Add(id);
                frontier = next;
            }

            var bfsSurvivors = visited.Where(allNodesById.ContainsKey).ToHashSet(StringComparer.Ordinal);
            var (bfsNodes, bfsEdges) = BuildOutput(bfsSurvivors, allNodesById, allEdges);
            return Cap(new FilteredCallgraph(bfsNodes, bfsEdges), options.MaxNodes);
        }

        if (moduleFilter is not null)
        {
            var moduleProcIds = allNodesById.Values
                .Where(n => n.Module == moduleFilter)
                .Select(n => n.Id)
                .ToHashSet(StringComparer.Ordinal);
            var survivingIds = new HashSet<string>(moduleProcIds, StringComparer.Ordinal);
            foreach (var e in allEdges)
            {
                if (moduleProcIds.Contains(e.From) && allNodesById.ContainsKey(e.To))
                    survivingIds.Add(e.To);
                if (moduleProcIds.Contains(e.To) && allNodesById.ContainsKey(e.From))
                    survivingIds.Add(e.From);
            }

            var (moduleNodes, moduleEdges) = BuildOutput(survivingIds, allNodesById, allEdges);
            return Cap(new FilteredCallgraph(moduleNodes, moduleEdges), options.MaxNodes);
        }

        var allProcIds = allNodesById.Keys.ToHashSet(StringComparer.Ordinal);
        var (allNodes, allEdgesOut) = BuildOutput(allProcIds, allNodesById, allEdges);
        return Cap(new FilteredCallgraph(allNodes, allEdgesOut), options.MaxNodes);
    }

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

        var degree = new Dictionary<string, int>(StringComparer.Ordinal);
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

    private static FilteredCallgraph Cap(FilteredCallgraph graph, int maxNodes)
    {
        if (graph.Nodes.Count > maxNodes)
            throw ToolError.GraphTooLarge(graph.Nodes.Count, maxNodes,
                "Add moduleName, add procedureName, or reduce depth.");
        return graph;
    }
}
