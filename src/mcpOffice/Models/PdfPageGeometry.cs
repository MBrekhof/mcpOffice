namespace McpOffice.Models;

/// <summary>
/// Page geometry. Width/Height are PDF points (1/72 inch) taken from the crop box.
/// Named ...Geometry rather than ...Info because DevExpress already has a PdfPageInfo.
/// </summary>
public sealed record PdfPageGeometry(
    int PageNumber,
    double Width,
    double Height,
    int Rotation);
