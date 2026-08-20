namespace McpOffice.Models;

/// <summary>
/// Result of pdf_read_text. <paramref name="Truncated"/> is true when maxChars stopped the read
/// before every requested page was returned.
/// </summary>
public sealed record PdfTextResult(
    int PageCount,
    IReadOnlyList<PdfPageText> Pages,
    int CharCount,
    bool Truncated);
