using System.Text;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba.Rendering;

public sealed class MermaidCallgraphRenderer : ICallgraphRenderer
{
    public string Render(FilteredCallgraph graph, CallgraphRenderOptions options)
    {
        var sb = new StringBuilder();
        sb.AppendLine("flowchart TD");

        if (options.Layout == "clustered")
            EmitClustered(sb, graph);
        else
            EmitFlat(sb, graph);

        EmitEdges(sb, graph);
        EmitClassDefs(sb);

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
            sb.Append("  subgraph ").AppendLine(MangleId(group.Key));
            foreach (var node in group)
            {
                sb.Append("    ");
                EmitNode(sb, node, useFqnLabel: false);
                sb.AppendLine();
            }
            sb.AppendLine("  end");
        }

        foreach (var ext in graph.Nodes.Where(n => n.IsExternal))
        {
            sb.Append("  ");
            EmitNode(sb, ext, useFqnLabel: false);
            sb.AppendLine();
        }
    }

    private static void EmitFlat(StringBuilder sb, FilteredCallgraph graph)
    {
        foreach (var node in graph.Nodes)
        {
            sb.Append("  ");
            EmitNode(sb, node, useFqnLabel: !node.IsExternal);
            sb.AppendLine();
        }
    }

    private static void EmitNode(StringBuilder sb, CallgraphNode node, bool useFqnLabel)
    {
        var id = MangleId(node.Id);
        var label = EscapeLabel(useFqnLabel ? node.Id : node.Label);

        if (node.IsEventHandler)
            sb.Append(id).Append("([").Append(label).Append("])");
        else
            sb.Append(id).Append('[').Append(label).Append(']');

        if (node.IsExternal) sb.Append(":::external");
        else if (node.IsEventHandler) sb.Append(":::handler");
        else if (node.IsOrphan) sb.Append(":::orphan");
    }

    private static void EmitEdges(StringBuilder sb, FilteredCallgraph graph)
    {
        foreach (var e in graph.Edges)
        {
            sb.Append("  ").Append(MangleId(e.FromId));
            sb.Append(e.Resolved ? " --> " : " -.-> ");
            sb.AppendLine(MangleId(e.ToId));
        }
    }

    private static void EmitClassDefs(StringBuilder sb)
    {
        sb.AppendLine("  classDef handler fill:#e1f5ff,stroke:#0277bd");
        sb.AppendLine("  classDef orphan stroke-dasharray:5 5");
        sb.AppendLine("  classDef external fill:#f5f5f5,stroke-dasharray:3 3");
    }

    private static string MangleId(string id)
    {
        var chars = id.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray();
        return new string(chars);
    }

    private static string EscapeLabel(string label)
    {
        return label
            .Replace("\"", "&quot;")
            .Replace("[", "&#91;")
            .Replace("]", "&#93;")
            .Replace("(", "&#40;")
            .Replace(")", "&#41;");
    }
}
