using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class CouplingComputerTests
{
    private static ExcelVbaCallEdge Edge(string from, string to, bool resolved = true) =>
        new(from, to, resolved,
            new ExcelVbaSiteRef(from.Split('.')[0], from.Split('.')[1], 1));

    private static IReadOnlyList<(string Name, string Kind)> Modules(params string[] names) =>
        names.Select(n => (n, "standard")).ToList();

    [Fact]
    public void ModuleCoupling_zeros_for_isolated_module()
    {
        var result = CouplingComputer.Compute(
            Modules("Lonely"),
            Array.Empty<ExcelVbaCallEdge>());
        var m = Assert.Single(result.Coupling);
        Assert.Equal("Lonely", m.Module);
        Assert.Equal(0, m.Ca);
        Assert.Equal(0, m.Ce);
        Assert.Equal(0.0, m.Instability);
        Assert.Equal(0, m.InternalEdges);
    }

    [Fact]
    public void ModuleCoupling_counts_external_edges_for_ca_and_ce()
    {
        // A → B (1 edge). A: Ce=1, Ca=0. B: Ca=1, Ce=0.
        var edges = new[] { Edge("A.f", "B.g") };
        var result = CouplingComputer.Compute(Modules("A", "B"), edges);

        var a = result.Coupling.Single(c => c.Module == "A");
        Assert.Equal(0, a.Ca);
        Assert.Equal(1, a.Ce);
        Assert.Equal(1.0, a.Instability);

        var b = result.Coupling.Single(c => c.Module == "B");
        Assert.Equal(1, b.Ca);
        Assert.Equal(0, b.Ce);
        Assert.Equal(0.0, b.Instability);
    }

    [Fact]
    public void ModuleCoupling_internalEdges_counts_intra_module_edges()
    {
        var edges = new[] { Edge("A.f", "A.g"), Edge("A.g", "A.h") };
        var result = CouplingComputer.Compute(Modules("A"), edges);
        var a = result.Coupling.Single();
        Assert.Equal(2, a.InternalEdges);
        Assert.Equal(0, a.Ca);
        Assert.Equal(0, a.Ce);
    }

    [Fact]
    public void ModuleCoupling_excludes_unresolved_edges()
    {
        var edges = new[] { Edge("A.f", "Unknown.x", resolved: false) };
        var result = CouplingComputer.Compute(Modules("A", "B"), edges);
        Assert.All(result.Coupling, m => Assert.Equal(0, m.Ce));
    }

    [Fact]
    public void ModuleCoupling_dedupes_repeated_edges()
    {
        // Same from/to across two call sites should count once.
        var edges = new[] { Edge("A.f", "B.g"), Edge("A.f", "B.g") };
        var result = CouplingComputer.Compute(Modules("A", "B"), edges);
        var a = result.Coupling.Single(c => c.Module == "A");
        Assert.Equal(1, a.Ce);
    }

    [Fact]
    public void ModuleCoupling_instability_balanced_module()
    {
        // C is called once and calls once → I = 1/(1+1) = 0.5.
        var edges = new[] { Edge("A.f", "C.x"), Edge("C.x", "B.g") };
        var result = CouplingComputer.Compute(Modules("A", "B", "C"), edges);
        var c = result.Coupling.Single(m => m.Module == "C");
        Assert.Equal(1, c.Ca);
        Assert.Equal(1, c.Ce);
        Assert.Equal(0.5, c.Instability);
    }

    [Fact]
    public void Pairs_emit_directional_edge_counts_sorted_desc()
    {
        var edges = new[]
        {
            Edge("A.f", "B.g"),
            Edge("A.f", "B.h"),
            Edge("A.k", "B.g"),
            Edge("B.g", "A.f")
        };
        var result = CouplingComputer.Compute(Modules("A", "B"), edges);

        Assert.Equal(2, result.Pairs.Count);
        var ab = result.Pairs.First();
        Assert.Equal("A", ab.From);
        Assert.Equal("B", ab.To);
        Assert.Equal(3, ab.EdgeCount);

        var ba = result.Pairs.Last();
        Assert.Equal("B", ba.From);
        Assert.Equal("A", ba.To);
        Assert.Equal(1, ba.EdgeCount);
    }

    [Fact]
    public void Pairs_omit_zero_count_pairs()
    {
        var result = CouplingComputer.Compute(Modules("A", "B", "C"),
            new[] { Edge("A.f", "B.g") });
        // Only A→B is non-zero.
        Assert.Single(result.Pairs);
        Assert.Equal("A", result.Pairs[0].From);
        Assert.Equal("B", result.Pairs[0].To);
    }

    [Fact]
    public void Pairs_stable_sort_alphabetical_within_same_count()
    {
        var edges = new[]
        {
            Edge("Z.f", "A.g"),
            Edge("M.f", "B.g")
        };
        var result = CouplingComputer.Compute(Modules("A", "B", "M", "Z"), edges);
        Assert.Equal(2, result.Pairs.Count);
        // Both have edgeCount=1; alphabetical by From: M before Z.
        Assert.Equal("M", result.Pairs[0].From);
        Assert.Equal("Z", result.Pairs[1].From);
    }
}
