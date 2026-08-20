namespace McpOffice.Models;

public sealed record PdfExtractImagesResult(
    string OutputDirectory,
    int ImageCount,
    IReadOnlyList<PdfExtractedImage> Images,
    bool Truncated);
