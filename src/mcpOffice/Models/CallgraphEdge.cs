namespace McpOffice.Models;

public sealed record CallgraphEdge(
    string FromId,
    string ToId,
    bool Resolved);
