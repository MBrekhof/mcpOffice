namespace McpOffice.Models;

/// <summary>
/// One word and where it sits on the page. X/Y are PDF points with the origin at the
/// TOP-left of the page (Y grows downward) — flipped from PDF's native bottom-left origin
/// so that sorting by Y ascending gives reading order.
/// </summary>
public sealed record PdfWordBox(
    string Text,
    double X,
    double Y,
    double Width,
    double Height,
    double? FontSize,
    string? FontName);
