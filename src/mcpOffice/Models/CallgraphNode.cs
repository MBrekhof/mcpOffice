namespace McpOffice.Models;

public sealed record CallgraphNode(
    string Id,
    string Label,
    string? Module,
    bool IsEventHandler,
    bool IsOrphan,
    bool IsExternal);
