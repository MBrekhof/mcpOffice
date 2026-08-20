namespace McpOffice.Models;

public sealed record PdfSearchResult(
    string Query,
    int HitCount,
    IReadOnlyList<PdfSearchHit> Hits,
    bool Truncated);
