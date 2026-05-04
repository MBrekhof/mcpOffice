using McpOffice.Models;
using McpOffice.Services.Excel.Vba.Rendering;

namespace McpOffice.Tests.Excel.Vba.Rendering;

public class MermaidCallgraphRendererTests
{
    private static readonly MermaidCallgraphRenderer R = new();

    [Fact]
    public void Empty_graph_emits_flowchart_header()
    {
        var output = R.Render(
            new FilteredCallgraph(Array.Empty<CallgraphNode>(), Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions());

        Assert.StartsWith("flowchart TD", output);
    }

    [Fact]
    public void Single_node_clustered_wraps_in_subgraph()
    {
        var node = new CallgraphNode("M.P1", "P1", "M", IsEventHandler: false, IsOrphan: true, IsExternal: false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "clustered"));

        Assert.Contains("subgraph M", output);
        Assert.Contains("end", output);
        Assert.Contains("M_P1", output);
        Assert.Contains("[P1]", output);
    }

    [Fact]
    public void Single_node_flat_no_subgraphs()
    {
        var node = new CallgraphNode("M.P1", "P1", "M", false, true, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.DoesNotContain("subgraph", output);
        Assert.Contains("M_P1", output);
        Assert.Contains("[M.P1]", output);
    }

    [Fact]
    public void Edge_resolved_emits_solid_arrow()
    {
        var p1 = new CallgraphNode("M.P1", "P1", "M", false, false, false);
        var p2 = new CallgraphNode("M.P2", "P2", "M", false, false, false);
        var edge = new CallgraphEdge("M.P1", "M.P2", Resolved: true);

        var output = R.Render(
            new FilteredCallgraph(new[] { p1, p2 }, new[] { edge }),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Matches(@"M_P1\s*-->\s*M_P2", output);
    }

    [Fact]
    public void Edge_unresolved_emits_dashed_arrow()
    {
        var p1 = new CallgraphNode("M.P1", "P1", "M", false, false, false);
        var ext = new CallgraphNode("__ext__MsgBox", "MsgBox", null, false, false, true);
        var edge = new CallgraphEdge("M.P1", "__ext__MsgBox", Resolved: false);

        var output = R.Render(
            new FilteredCallgraph(new[] { p1, ext }, new[] { edge }),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Matches(@"M_P1\s*-\.->\s*__ext__MsgBox", output);
    }

    [Fact]
    public void Procedure_name_with_brackets_is_escaped_in_label()
    {
        var node = new CallgraphNode("M.[Bracketed Name]", "[Bracketed Name]", "M", false, false, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.DoesNotContain("[[Bracketed Name]]", output);
        Assert.Contains("&#91;Bracketed Name&#93;", output);
    }

    [Fact]
    public void Procedure_name_with_parens_is_escaped_in_handler_node()
    {
        var node = new CallgraphNode("M.Foo(bar)", "Foo(bar)", "M",
            IsEventHandler: true, IsOrphan: false, IsExternal: false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Contains("Foo&#40;bar&#41;", output);
    }

    [Fact]
    public void Module_name_with_space_is_mangled_in_subgraph_id()
    {
        var node = new CallgraphNode("Sheet 1.P1", "P1", "Sheet 1", false, false, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "clustered"));

        Assert.Contains("subgraph Sheet_1", output);
    }

    [Fact]
    public void Subgraph_open_count_matches_end_count()
    {
        var nodes = new[]
        {
            new CallgraphNode("M1.A", "A", "M1", false, false, false),
            new CallgraphNode("M1.B", "B", "M1", false, false, false),
            new CallgraphNode("M2.C", "C", "M2", false, false, false),
            new CallgraphNode("__ext__MsgBox", "MsgBox", null, false, false, true),
        };
        var output = R.Render(
            new FilteredCallgraph(nodes, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "clustered"));

        var subgraphCount = System.Text.RegularExpressions.Regex.Matches(output, @"^\s*subgraph\b", System.Text.RegularExpressions.RegexOptions.Multiline).Count;
        var endCount = System.Text.RegularExpressions.Regex.Matches(output, @"^\s*end\s*$", System.Text.RegularExpressions.RegexOptions.Multiline).Count;
        Assert.Equal(subgraphCount, endCount);
        Assert.Equal(2, subgraphCount);
    }
}
