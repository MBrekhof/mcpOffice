using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaCallGraphReachabilityTests
{
    private static ExcelVbaCallEdge Edge(string from, string to, bool resolved = true) =>
        new(from, to, resolved, new ExcelVbaSiteRef("M", "P", 1));

    [Fact]
    public void Chain_is_fully_reachable()
    {
        var r = VbaCallGraphReachability.Compute(["M.A", "M.B", "M.C"], ["M.A"], [Edge("M.A", "M.B"), Edge("M.B", "M.C")]);

        Assert.Equal(["M.A", "M.B", "M.C"], r.Reachable.Order());
        Assert.Empty(r.Unreachable);
    }

    [Fact]
    public void Diamond_visits_shared_node_once()
    {
        var r = VbaCallGraphReachability.Compute(["M.A", "M.B", "M.C", "M.D"], ["M.A"],
            [Edge("M.A", "M.B"), Edge("M.A", "M.C"), Edge("M.B", "M.D"), Edge("M.C", "M.D")]);

        Assert.Equal(4, r.Reachable.Count);
        Assert.Empty(r.Unreachable);
    }

    [Fact]
    public void Cycle_terminates_and_isolated_node_is_unreachable()
    {
        var r = VbaCallGraphReachability.Compute(["M.A", "M.B", "M.X"], ["M.A"], [Edge("M.A", "M.B"), Edge("M.B", "M.A")]);

        Assert.Equal(["M.A", "M.B"], r.Reachable.Order());
        Assert.Equal(["M.X"], r.Unreachable);
    }

    [Fact]
    public void Unresolved_edges_are_ignored()
    {
        var r = VbaCallGraphReachability.Compute(["M.A", "M.B"], ["M.A"], [Edge("M.A", "M.B", resolved: false)]);

        Assert.Equal(["M.B"], r.Unreachable);
    }

    [Fact]
    public void Extra_edges_rescue_otherwise_unreachable_procedures()
    {
        var r = VbaCallGraphReachability.Compute(["M.A", "M.B"], ["M.A"], [], extraEdges: [("M.A", "M.B")]);

        Assert.Empty(r.Unreachable);
    }

    [Fact]
    public void Matching_is_case_insensitive_and_reachable_keeps_declared_casing()
    {
        var r = VbaCallGraphReachability.Compute(["Mod.Foo", "Mod.Bar"], ["mod.foo"], [Edge("MOD.FOO", "mod.BAR")]);

        Assert.Equal(["Mod.Bar", "Mod.Foo"], r.Reachable.Order());
        Assert.Contains("MOD.BAR", r.Reachable);
        Assert.Empty(r.Unreachable);
    }

    [Fact]
    public void No_entry_points_means_everything_unreachable_sorted_ignoring_case()
    {
        var r = VbaCallGraphReachability.Compute(["b.Z", "A.y", "a.X"], [], [Edge("b.Z", "A.y")]);

        Assert.Empty(r.Reachable);
        Assert.Equal(["a.X", "A.y", "b.Z"], r.Unreachable);
    }

    [Fact]
    public void Unknown_entry_points_and_edge_targets_are_ignored()
    {
        var r = VbaCallGraphReachability.Compute(["M.A"], ["M.Ghost", "M.A"], [Edge("M.A", "M.Missing")]);

        Assert.Equal(["M.A"], r.Reachable);
        Assert.Empty(r.Unreachable);
    }
}
