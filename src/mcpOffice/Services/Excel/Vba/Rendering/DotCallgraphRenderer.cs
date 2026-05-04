using System.Text;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba.Rendering;

public sealed class DotCallgraphRenderer : ICallgraphRenderer
{
    public string Render(FilteredCallgraph graph, CallgraphRenderOptions options)
    {
        var sb = new StringBuilder();
        sb.Append("digraph G {\n");
        sb.Append("  rankdir=TB;\n");
        sb.Append("  node [shape=box];\n");

        if (options.Layout == "clustered")
            EmitClustered(sb, graph);
        else
            EmitFlat(sb, graph);

        EmitEdges(sb, graph);
        sb.Append("}\n");
        return sb.ToString();
    }

    private static void EmitClustered(StringBuilder sb, FilteredCallgraph graph)
    {
        var grouped = graph.Nodes
            .Where(n => !n.IsExternal)
            .GroupBy(n => n.Module!)
            .OrderBy(g => g.Key, StringComparer.Ordinal);

        foreach (var group in grouped)
        {
            var clusterId = "cluster_" + Mangle(group.Key);
            sb.Append("  subgraph ").Append(clusterId).Append(" {\n");
            sb.Append("    label=").Append(Quote(group.Key)).Append(";\n");
            foreach (var node in group)
            {
                sb.Append("    ");
                EmitNode(sb, node, useFqnLabel: false);
                sb.Append('\n');
            }
            sb.Append("  }\n");
        }

        foreach (var ext in graph.Nodes.Where(n => n.IsExternal))
        {
            sb.Append("  ");
            EmitNode(sb, ext, useFqnLabel: false);
            sb.Append('\n');
        }
    }

    private static void EmitFlat(StringBuilder sb, FilteredCallgraph graph)
    {
        foreach (var node in graph.Nodes)
        {
            sb.Append("  ");
            EmitNode(sb, node, useFqnLabel: !node.IsExternal);
            sb.Append('\n');
        }
    }

    private static void EmitNode(StringBuilder sb, CallgraphNode node, bool useFqnLabel)
    {
        var id = Quote(node.Id);
        var label = useFqnLabel ? node.Id : node.Label;

        var attrs = new List<string> { $"label={Quote(label)}" };

        if (node.IsExternal)
        {
            attrs.Add("shape=box");
            attrs.Add("style=\"dashed,filled\"");
            attrs.Add("fillcolor=\"#f5f5f5\"");
        }
        else if (node.IsEventHandler)
        {
            attrs.Add("shape=oval");
            attrs.Add("style=\"filled\"");
            attrs.Add("fillcolor=\"#e1f5ff\"");
        }
        else if (node.IsOrphan)
        {
            attrs.Add("style=\"dashed\"");
        }

        sb.Append(id).Append(" [").Append(string.Join(", ", attrs)).Append("];");
    }

    private static void EmitEdges(StringBuilder sb, FilteredCallgraph graph)
    {
        foreach (var e in graph.Edges)
        {
            sb.Append("  ").Append(Quote(e.FromId)).Append(" -> ").Append(Quote(e.ToId));
            if (!e.Resolved)
                sb.Append(" [style=\"dashed\"]");
            sb.Append(";\n");
        }
    }

    private static string Mangle(string s)
    {
        var chars = s.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray();
        return new string(chars);
    }

    private static string Quote(string s) => "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
