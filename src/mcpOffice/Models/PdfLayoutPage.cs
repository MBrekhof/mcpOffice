namespace McpOffice.Models;

/// <summary>
/// One page of positioned text. Exactly one of <paramref name="Words"/> / <paramref name="Lines"/>
/// is populated, depending on the requested granularity.
/// </summary>
public sealed record PdfLayoutPage(
    int PageNumber,
    double Width,
    double Height,
    IReadOnlyList<PdfWordBox>? Words,
    IReadOnlyList<PdfTextLine>? Lines);
