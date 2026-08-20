namespace McpOffice.Models;

public sealed record PdfLayoutResult(
    int PageCount,
    string Granularity,
    int WordCount,
    IReadOnlyList<PdfLayoutPage> Pages,
    bool Truncated);
