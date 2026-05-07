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
