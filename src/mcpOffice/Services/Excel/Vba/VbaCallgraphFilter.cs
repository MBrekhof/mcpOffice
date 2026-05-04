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

        string? moduleFilter = null;
        if (!string.IsNullOrWhiteSpace(options.ModuleName))
        {
            var match = analysis.Modules.FirstOrDefault(m =>
                string.Equals(m.Name, options.ModuleName, StringComparison.OrdinalIgnoreCase));
            if (match is null)
                throw ToolError.ModuleNotFound(options.ModuleName, analysis.Modules.Select(m => m.Name));
            moduleFilter = match.Name;
        }

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
                var fromInModule = moduleProcIds.Contains(e.From);
                var toInModule = moduleProcIds.Contains(e.To);
                if (fromInModule && allNodesById.ContainsKey(e.To))
                    survivingIds.Add(e.To);
                if (toInModule && allNodesById.ContainsKey(e.From))
                    survivingIds.Add(e.From);
            }

            // Iterate the dictionary so node order follows declaration order (deterministic for renderers).
            var moduleNodes = allNodesById.Values.Where(n => survivingIds.Contains(n.Id)).ToList();
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
}
