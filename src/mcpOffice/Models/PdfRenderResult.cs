namespace McpOffice.Models;

public sealed record PdfRenderResult(
    string OutputPath,
    int PageNumber,
    int PixelWidth,
    int PixelHeight,
    long BytesWritten);
