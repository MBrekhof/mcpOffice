namespace McpOffice.Models;

public sealed record FilteredCallgraph(
    IReadOnlyList<CallgraphNode> Nodes,
    IReadOnlyList<CallgraphEdge> Edges);
