namespace McpOffice.Models;

/// <summary>
/// A bookmark. <paramref name="PageNumber"/> is null when the bookmark has no page destination
/// (an action-only bookmark, e.g. a URI link).
/// </summary>
public sealed record PdfOutlineNode(
    string Title,
    int Level,
    int? PageNumber,
    IReadOnlyList<PdfOutlineNode> Children);
