namespace McpOffice.Models;

/// <summary>A single match. Coordinates follow <see cref="PdfWordBox"/> (origin top-left).</summary>
public sealed record PdfSearchHit(
    int PageNumber,
    string Text,
    double X,
    double Y,
    double Width,
    double Height);
