using System.ComponentModel;
using McpOffice.Services.Pdf;
using ModelContextProtocol.Server;

namespace McpOffice.Tools;

[McpServerToolType]
public static class PdfTools
{
    private static readonly IPdfDocumentService Service = new PdfDocumentService();

    [McpServerTool(Name = "pdf_get_metadata")]
    [Description("Returns PDF document properties (title, author, subject, keywords, creator, producer, created, modified), pdf version, permission flags (allowDataExtraction/allowPrinting/allowModifying), bookmarkCount, pageCount, and a per-page array of {pageNumber, width, height, rotation} in PDF points (1/72 inch).")]
    public static object PdfGetMetadata(
        [Description("Absolute path to the .pdf file")] string path)
        => Service.GetMetadata(path);

    [McpServerTool(Name = "pdf_read_text")]
    [Description("Extracts text per page. Set preserveLayout=true to rebuild the page as a fixed-width grid (like pdftotext -layout) so column-based reports stay aligned and sliceable by character position; leave it false for prose, which is faster. Returns {pageCount, pages:[{pageNumber, text, charCount}], charCount, truncated}. maxChars caps the total and sets truncated.")]
    public static object PdfReadText(
        [Description("Absolute path to the .pdf file")] string path,
        [Description("Pages to read: '1', '2-5', '1,3,7-9', '5-' (to the end). Omit for every page.")] string? pageRange = null,
        [Description("Rebuild the fixed-width column layout instead of returning content-stream order. Use for tabular/monospaced reports.")] bool preserveLayout = false,
        [Description("Maximum characters to return across all pages. Prevents accidental huge reads.")] int maxChars = 200000)
        => Service.ReadText(path, pageRange, preserveLayout, maxChars);

    [McpServerTool(Name = "pdf_read_layout")]
    [Description("Returns positioned text: every word (or visual line) with x/y/width/height in PDF points, origin at the TOP-left of the page so sorting by y ascending is reading order. Use this when you need to know which column a value sits in - PDF itself stores no lines or columns, so granularity='line' groups words by vertical position. maxWords caps the walk and sets truncated.")]
    public static object PdfReadLayout(
        [Description("Absolute path to the .pdf file")] string path,
        [Description("Pages to read: '1', '2-5', '1,3,7-9', '5-' (to the end). Omit for every page.")] string? pageRange = null,
        [Description("'line' (default) groups words into visual lines; 'word' returns each word separately.")] string granularity = "line",
        [Description("Include fontSize and fontName on each word. Default false.")] bool includeFontInfo = false,
        [Description("Maximum words to walk. Prevents accidental huge reads on long documents.")] int maxWords = 50000)
        => Service.ReadLayout(path, pageRange, granularity, includeFontInfo, maxWords);

    [McpServerTool(Name = "pdf_find_text")]
    [Description("Searches the document and returns each match with its page and bounding box {pageNumber, text, x, y, width, height} (origin top-left). Returns {query, hitCount, hits, truncated}.")]
    public static object PdfFindText(
        [Description("Absolute path to the .pdf file")] string path,
        [Description("Text to search for")] string query,
        [Description("Match case. Default false.")] bool caseSensitive = false,
        [Description("Match whole words only. Default false.")] bool wholeWords = false,
        [Description("Maximum matches to return.")] int maxResults = 500)
        => Service.FindText(path, query, caseSensitive, wholeWords, maxResults);

    [McpServerTool(Name = "pdf_render_page")]
    [Description("Renders one page to an image file so it can be viewed. Format is inferred from the outputPath extension when omitted (png, jpg/jpeg, bmp, gif, tiff). Returns {outputPath, pageNumber, pixelWidth, pixelHeight, bytesWritten}. Use this for scanned PDFs, or to check what a page actually looks like when the extracted text is ambiguous.")]
    public static object PdfRenderPage(
        [Description("Absolute path to the .pdf file")] string path,
        [Description("1-based page number to render")] int pageNumber,
        [Description("Absolute path of the image file to write. Parent directory is created if missing.")] string outputPath,
        [Description("Render resolution in DPI (12-1200). 150 is readable; 300 for fine print.")] int dpi = 150,
        [Description("Optional output format override: png, jpg, jpeg, bmp, gif, tiff.")] string? format = null,
        [Description("Overwrite outputPath if it already exists. Defaults to false.")] bool overwrite = false)
        => Service.RenderPage(path, pageNumber, outputPath, dpi, format, overwrite);

    [McpServerTool(Name = "pdf_extract_images")]
    [Description("Writes the raster images embedded in the PDF to a directory as PNG files, named page{NNN}_img{NNN}.png. Each entry reports its placement on the page (x/y/width/height in points, origin top-left) and its stored pixel size. Use minPixelSize to skip rules, bullets and other tiny decorations.")]
    public static object PdfExtractImages(
        [Description("Absolute path to the .pdf file")] string path,
        [Description("Absolute path to the directory to write images into. Created if missing.")] string outputDirectory,
        [Description("Pages to scan: '1', '2-5', '1,3,7-9', '5-' (to the end). Omit for every page.")] string? pageRange = null,
        [Description("Skip images narrower or shorter than this many pixels. Default 16.")] int minPixelSize = 16,
        [Description("Maximum images to extract.")] int maxImages = 200,
        [Description("Overwrite existing files in the output directory. Defaults to false.")] bool overwrite = false)
        => Service.ExtractImages(path, outputDirectory, pageRange, minPixelSize, maxImages, overwrite);

    [McpServerTool(Name = "pdf_get_outline")]
    [Description("Returns the bookmark tree as nested {title, level, pageNumber, children}. pageNumber is null for bookmarks with no page destination. Empty array when the PDF has no bookmarks.")]
    public static object PdfGetOutline(
        [Description("Absolute path to the .pdf file")] string path)
        => Service.GetOutline(path);
}
