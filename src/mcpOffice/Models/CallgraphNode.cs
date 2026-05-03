namespace McpOffice.Models;

/// <summary>
/// A node in a filtered call graph. <see cref="Id"/> is the canonical FQN
/// (e.g., "mdlScreeningDB.ReadExports") for resolved procedures, or "__ext__&lt;name&gt;"
/// for unresolved/external callees. Renderers are responsible for mangling Id
/// into format-safe identifiers.
/// </summary>
public sealed record CallgraphNode(
    string Id,
    string Label,
    string? Module,             // null for external nodes
    bool IsEventHandler,
    bool IsOrphan,
    bool IsExternal);
