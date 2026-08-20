namespace McpOffice.Models;

/// <summary>
/// One embedded raster image written to disk. X/Y/Width/Height are the image's placement on the
/// page in PDF points with a top-left origin (see <see cref="PdfWordBox"/>); PixelWidth/PixelHeight
/// are the stored resolution.
/// </summary>
public sealed record PdfExtractedImage(
    int PageNumber,
    string OutputPath,
    double X,
    double Y,
    double Width,
    double Height,
    int PixelWidth,
    int PixelHeight,
    long BytesWritten);
