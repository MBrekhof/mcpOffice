using McpOffice.Models;
using McpOffice.Services.Excel.Vba.Rendering;

namespace McpOffice.Tests.Excel.Vba.Rendering;

public class DotCallgraphRendererTests
{
    private static readonly DotCallgraphRenderer R = new();

    [Fact]
    public void Empty_graph_emits_digraph_header_and_braces()
    {
        var output = R.Render(
            new FilteredCallgraph(Array.Empty<CallgraphNode>(), Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions());

        Assert.StartsWith("digraph G {", output);
        Assert.EndsWith("}\n", output.Replace("\r\n", "\n"));
    }

    [Fact]
    public void Single_node_clustered_wraps_in_subgraph_cluster()
    {
        var node = new CallgraphNode("M.P1", "P1", "M", false, true, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "clustered"));

        Assert.Contains("subgraph cluster_M", output);
        Assert.Contains("\"M.P1\"", output);
        Assert.Contains("label=\"P1\"", output);
    }

    [Fact]
    public void Flat_uses_FQN_label()
    {
        var node = new CallgraphNode("M.P1", "P1", "M", false, false, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.DoesNotContain("subgraph", output);
        Assert.Contains("label=\"M.P1\"", output);
    }

    [Fact]
    public void Resolved_edge_solid_unresolved_dashed()
    {
        var p1 = new CallgraphNode("M.P1", "P1", "M", false, false, false);
        var p2 = new CallgraphNode("M.P2", "P2", "M", false, false, false);
        var ext = new CallgraphNode("__ext__MsgBox", "MsgBox", null, false, false, true);

        var output = R.Render(
            new FilteredCallgraph(
                new[] { p1, p2, ext },
                new[]
                {
                    new CallgraphEdge("M.P1", "M.P2", Resolved: true),
                    new CallgraphEdge("M.P1", "__ext__MsgBox", Resolved: false),
                }),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Matches("\"M.P1\"\\s*->\\s*\"M.P2\"", output);
        Assert.Contains("style=\"dashed\"", output);
    }

    [Fact]
    public void Handler_node_uses_oval_shape()
    {
        var node = new CallgraphNode("M.Open", "Open", "M",
            IsEventHandler: true, IsOrphan: false, IsExternal: false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Contains("shape=oval", output);
    }

    [Fact]
    public void External_node_styled_dashed()
    {
        var ext = new CallgraphNode("__ext__MsgBox", "MsgBox", null, false, false, true);
        var output = R.Render(
            new FilteredCallgraph(new[] { ext }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Contains("style=\"dashed,filled\"", output);
        Assert.Contains("fillcolor=\"#f5f5f5\"", output);
    }

    [Fact]
    public void Procedure_id_with_quote_is_escaped()
    {
        var node = new CallgraphNode("M.It\"s", "It\"s", "M", false, false, false);
        var output = R.Render(
            new FilteredCallgraph(new[] { node }, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "flat"));

        Assert.Contains("\\\"", output);
    }

    [Fact]
    public void Brace_balance_holds()
    {
        var nodes = new[]
        {
            new CallgraphNode("M1.A", "A", "M1", false, false, false),
            new CallgraphNode("M2.B", "B", "M2", false, false, false),
        };
        var output = R.Render(
            new FilteredCallgraph(nodes, Array.Empty<CallgraphEdge>()),
            new CallgraphRenderOptions(Layout: "clustered"));

        Assert.Equal(output.Count(c => c == '{'), output.Count(c => c == '}'));
    }
}
