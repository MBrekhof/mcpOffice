namespace McpOffice.Models;

public sealed record PdfDocumentMetadata(
    int PageCount,
    string? Title,
    string? Author,
    string? Subject,
    string? Keywords,
    string? Creator,
    string? Producer,
    DateTimeOffset? Created,
    DateTimeOffset? Modified,
    string Version,
    bool AllowDataExtraction,
    bool AllowPrinting,
    bool AllowModifying,
    int BookmarkCount,
    IReadOnlyList<PdfPageGeometry> Pages);
