using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

/// <summary>BFS over the FQN call graph. VBA names are case-insensitive; results keep the declared casing.</summary>
internal static class VbaCallGraphReachability
{
    public sealed record Result(IReadOnlySet<string> Reachable, IReadOnlyList<string> Unreachable);

    public static Result Compute(
        IEnumerable<string> allProcedureFqns,
        IEnumerable<string> entryPointFqns,
        IEnumerable<ExcelVbaCallEdge> edges,
        IEnumerable<(string From, string To)>? extraEdges = null)
    {
        var ci = StringComparer.OrdinalIgnoreCase;
        var all = new HashSet<string>(allProcedureFqns, ci);

        var adjacency = new Dictionary<string, List<string>>(ci);
        foreach (var e in edges) if (e.Resolved) Add(e.From, e.To);
        foreach (var (from, to) in extraEdges ?? []) Add(from, to);

        var reachable = new HashSet<string>(ci);
        var queue = new Queue<string>();
        foreach (var entry in entryPointFqns) Visit(entry);
        while (queue.TryDequeue(out var current))
            if (adjacency.TryGetValue(current, out var targets))
                foreach (var t in targets) Visit(t);

        return new Result(reachable, all.Where(p => !reachable.Contains(p)).Order(ci).ToList());

        void Add(string from, string to)
        {
            if (!adjacency.TryGetValue(from, out var list)) adjacency[from] = list = [];
            list.Add(to);
        }

        // Unknown FQNs (external or unparsed targets) are dropped; known ones enter with declared casing.
        void Visit(string fqn)
        {
            if (all.TryGetValue(fqn, out var declared) && reachable.Add(declared)) queue.Enqueue(declared);
        }
    }
}
