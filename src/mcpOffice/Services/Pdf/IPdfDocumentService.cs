using McpOffice.Models;

namespace McpOffice.Services.Pdf;

public interface IPdfDocumentService
{
    PdfDocumentMetadata GetMetadata(string path);
    PdfTextResult ReadText(string path, string? pageRange, bool preserveLayout, int maxChars);
    PdfLayoutResult ReadLayout(string path, string? pageRange, string granularity, bool includeFontInfo, int maxWords);
    PdfSearchResult FindText(string path, string query, bool caseSensitive, bool wholeWords, int maxResults);
    PdfRenderResult RenderPage(string path, int pageNumber, string outputPath, int dpi, string? format, bool overwrite);
    PdfExtractImagesResult ExtractImages(string path, string outputDirectory, string? pageRange, int minPixelSize, int maxImages, bool overwrite);
    IReadOnlyList<PdfOutlineNode> GetOutline(string path);
}
