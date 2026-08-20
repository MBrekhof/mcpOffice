namespace McpOffice.Models;

/// <summary>
/// Words grouped into a visual line, ordered left to right. X is the left edge of the first
/// word, Y the line's distance from the top of the page (see <see cref="PdfWordBox"/>).
/// </summary>
public sealed record PdfTextLine(
    int PageNumber,
    double X,
    double Y,
    string Text,
    IReadOnlyList<PdfWordBox> Words);
