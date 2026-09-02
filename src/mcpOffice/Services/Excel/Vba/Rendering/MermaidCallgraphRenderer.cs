using System.Text;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba.Rendering;

// Renderer output is a wire format that lands in JSON-RPC payloads — line endings are LF only,
// not host-formatted text. Both renderers (Mermaid here, DOT in DotCallgraphRenderer) follow this.
public sealed class MermaidCallgraphRenderer : ICallgraphRenderer
{
    public string Render(FilteredCallgraph graph, CallgraphRenderOptions options)
    {
        var sb = new StringBuilder();
        sb.Append("flowchart TD\n");

        // Layout match is exact: validation of the user-supplied value is the tool layer's job (Task 15).
        // Anything that isn't literally "clustered" falls through to flat — the safe default.
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
            sb.Append("  subgraph ").Append(MangleId(group.Key)).Append('\n');
            foreach (var node in group)
            {
                sb.Append("    ");
                EmitNode(sb, node, useFqnLabel: false);
                sb.Append('\n');
            }
            sb.Append("  end\n");
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
        var id = MangleId(node.Id);
        var label = EscapeLabel(useFqnLabel ? node.Id : node.Label);

        // Shape priority differs from class priority: handlers get the rounded shape, externals
        // get the :::external class. Per BuildOutput, externals never have IsEventHandler=true,
        // so the shape/class disagreement on a single node is unreachable.
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
            sb.Append(MangleId(e.ToId)).Append('\n');
        }
    }

    private static void EmitClassDefs(StringBuilder sb)
    {
        sb.Append("  classDef handler fill:#e1f5ff,stroke:#0277bd\n");
        sb.Append("  classDef orphan stroke-dasharray:5 5\n");
        sb.Append("  classDef external fill:#f5f5f5,stroke-dasharray:3 3\n");
    }

    private static string MangleId(string id)
    {
        // Mermaid IDs accept [A-Za-z0-9_] only. We replace everything else with '_'.
        // Theoretical collision: a module literally named "M_P1" mangles to the same id as
        // FQN "M.P1". Real workbooks haven't hit this; revisit with a bijective mapping if needed.
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
