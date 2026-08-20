using DevExpress.Drawing;
using DevExpress.Pdf;
using McpOffice.Models;
using ModelContextProtocol;

namespace McpOffice.Services.Pdf;

public sealed class PdfDocumentService : IPdfDocumentService
{
    public PdfDocumentMetadata GetMetadata(string path)
    {
        PathGuard.RequireExists(path);

        return Read(path, processor =>
        {
            var document = processor.Document;

            var pages = document.Pages
                .Select((page, index) => new PdfPageGeometry(
                    index + 1,
                    Math.Round(page.CropBox.Width, 2),
                    Math.Round(page.CropBox.Height, 2),
                    page.Rotate))
                .ToArray();

            return new PdfDocumentMetadata(
                document.Pages.Count,
                NullIfBlank(document.Title),
                NullIfBlank(document.Author),
                NullIfBlank(document.Subject),
                NullIfBlank(document.Keywords),
                NullIfBlank(document.Creator),
                NullIfBlank(document.Producer),
                document.CreationDate,
                document.ModDate,
                document.Version.ToString(),
                document.AllowDataExtraction,
                document.AllowPrinting,
                document.AllowModifying,
                CountBookmarks(document.Bookmarks),
                pages);
        });
    }

    public PdfTextResult ReadText(string path, string? pageRange, bool preserveLayout, int maxChars)
    {
        PathGuard.RequireExists(path);

        if (maxChars <= 0)
        {
            throw ToolError.RangeTooLarge("maxChars", maxChars, 1);
        }

        return Read(path, processor =>
        {
            var pageCount = processor.Document.Pages.Count;
            var wanted = PageRange.Parse(pageRange, pageCount);

            // Positioned words are only needed for the layout-preserving path; GetPageText is
            // markedly cheaper and is what most callers want.
            var wordsByPage = preserveLayout ? CollectWords(processor, wanted, includeFontInfo: false) : null;

            var pages = new List<PdfPageText>();
            var total = 0;
            var truncated = false;

            foreach (var pageNumber in wanted)
            {
                var text = preserveLayout
                    ? LayoutTextRenderer.Render(LineGrouper.Group(pageNumber, wordsByPage!.GetValueOrDefault(pageNumber, []) ))
                    : processor.GetPageText(pageNumber);

                text ??= string.Empty;

                if (total + text.Length > maxChars)
                {
                    var remaining = maxChars - total;
                    if (remaining > 0)
                    {
                        pages.Add(new PdfPageText(pageNumber, text[..remaining], remaining));
                        total += remaining;
                    }

                    truncated = true;
                    break;
                }

                pages.Add(new PdfPageText(pageNumber, text, text.Length));
                total += text.Length;
            }

            return new PdfTextResult(pageCount, pages, total, truncated);
        });
    }

    public PdfLayoutResult ReadLayout(string path, string? pageRange, string granularity, bool includeFontInfo, int maxWords)
    {
        PathGuard.RequireExists(path);

        var mode = (granularity ?? "line").Trim().ToLowerInvariant();
        if (mode is not ("word" or "line"))
        {
            throw ToolError.UnsupportedFormat($"granularity '{granularity}' (use 'word' or 'line')");
        }

        if (maxWords <= 0)
        {
            throw ToolError.RangeTooLarge("maxWords", maxWords, 1);
        }

        return Read(path, processor =>
        {
            var pageCount = processor.Document.Pages.Count;
            var wanted = PageRange.Parse(pageRange, pageCount);
            var wordsByPage = CollectWords(processor, wanted, includeFontInfo, maxWords, out var truncated);
            var wordCount = wordsByPage.Values.Sum(w => w.Count);

            var pages = wanted
                .Where(wordsByPage.ContainsKey)
                .Select(pageNumber =>
                {
                    var page = processor.Document.Pages[pageNumber - 1];
                    var words = wordsByPage[pageNumber];

                    return new PdfLayoutPage(
                        pageNumber,
                        Math.Round(page.CropBox.Width, 2),
                        Math.Round(page.CropBox.Height, 2),
                        mode == "word" ? words : null,
                        mode == "line" ? LineGrouper.Group(pageNumber, words) : null);
                })
                .ToArray();

            return new PdfLayoutResult(pageCount, mode, wordCount, pages, truncated);
        });
    }

    public PdfSearchResult FindText(string path, string query, bool caseSensitive, bool wholeWords, int maxResults)
    {
        PathGuard.RequireExists(path);

        if (string.IsNullOrEmpty(query))
        {
            throw ToolError.InvalidPageRange(query, "query must not be empty.");
        }

        if (maxResults <= 0)
        {
            throw ToolError.RangeTooLarge("maxResults", maxResults, 1);
        }

        return Read(path, processor =>
        {
            var parameters = new PdfTextSearchParameters
            {
                CaseSensitive = caseSensitive,
                WholeWords = wholeWords,
            };

            var hits = new List<PdfSearchHit>();
            var truncated = false;

            // FindText advances an internal cursor and reports Finished once it wraps.
            var results = processor.FindText(query, parameters);
            while (results.Status == PdfTextSearchStatus.Found)
            {
                var page = processor.Document.Pages[results.PageNumber - 1];

                foreach (var rectangle in results.Rectangles)
                {
                    if (hits.Count >= maxResults)
                    {
                        truncated = true;
                        break;
                    }

                    hits.Add(new PdfSearchHit(
                        results.PageNumber,
                        string.Join(' ', results.Words.Select(w => w.Text)),
                        Round(rectangle.Left),
                        Round(FlipY(rectangle.Top, page)),
                        Round(rectangle.Width),
                        Round(rectangle.Height)));
                }

                if (truncated)
                {
                    break;
                }

                results = processor.FindText(query, parameters);
            }

            return new PdfSearchResult(query, hits.Count, hits, truncated);
        });
    }

    public PdfRenderResult RenderPage(string path, int pageNumber, string outputPath, int dpi, string? format, bool overwrite)
    {
        PathGuard.RequireExists(path);
        PathGuard.RequireWritable(outputPath, overwrite);

        if (dpi is < 12 or > 1200)
        {
            throw ToolError.InvalidRenderOption($"dpi must be between 12 and 1200 (got {dpi}).");
        }

        var imageFormat = ResolveImageFormat(format, outputPath);

        return Read(path, processor =>
        {
            var pageCount = processor.Document.Pages.Count;
            if (pageNumber < 1 || pageNumber > pageCount)
            {
                throw ToolError.PageNotFound(pageNumber, pageCount);
            }

            using var bitmap = processor.CreateDXBitmap(pageNumber, PdfPageRenderingParameters.CreateWithResolution(dpi));
            bitmap.Save(outputPath, imageFormat);

            return new PdfRenderResult(
                outputPath,
                pageNumber,
                bitmap.Width,
                bitmap.Height,
                new FileInfo(outputPath).Length);
        });
    }

    public PdfExtractImagesResult ExtractImages(
        string path, string outputDirectory, string? pageRange, int minPixelSize, int maxImages, bool overwrite)
    {
        PathGuard.RequireExists(path);
        PathGuard.RequireAbsolute(outputDirectory);

        if (maxImages <= 0)
        {
            throw ToolError.RangeTooLarge("maxImages", maxImages, 1);
        }

        Directory.CreateDirectory(outputDirectory);

        return Read(path, processor =>
        {
            var pageCount = processor.Document.Pages.Count;
            var wanted = PageRange.Parse(pageRange, pageCount);
            var images = new List<PdfExtractedImage>();
            var truncated = false;

            foreach (var pageNumber in wanted)
            {
                var page = processor.Document.Pages[pageNumber - 1];
                var area = new PdfDocumentArea(pageNumber, page.CropBox);

                foreach (var box in processor.GetImagesInfo(area))
                {
                    using (box)
                    {
                        if (images.Count >= maxImages)
                        {
                            truncated = true;
                            break;
                        }

                        if (box.Bitmap.Width < minPixelSize || box.Bitmap.Height < minPixelSize)
                        {
                            continue;
                        }

                        var file = Path.Combine(
                            outputDirectory,
                            $"page{pageNumber:D3}_img{images.Count + 1:D3}.png");

                        if (File.Exists(file) && !overwrite)
                        {
                            throw ToolError.FileExists(file);
                        }

                        box.Bitmap.Save(file, DXImageFormat.Png);

                        images.Add(new PdfExtractedImage(
                            pageNumber,
                            file,
                            Round(box.Bounds.Left),
                            Round(FlipY(box.Bounds.Top, page)),
                            Round(box.Bounds.Width),
                            Round(box.Bounds.Height),
                            box.Bitmap.Width,
                            box.Bitmap.Height,
                            new FileInfo(file).Length));
                    }
                }

                if (truncated)
                {
                    break;
                }
            }

            return new PdfExtractImagesResult(outputDirectory, images.Count, images, truncated);
        });
    }

    public IReadOnlyList<PdfOutlineNode> GetOutline(string path)
    {
        PathGuard.RequireExists(path);

        return Read(path, processor => MapBookmarks(processor, processor.Document.Bookmarks, level: 1));
    }

    // ---- internals -------------------------------------------------------------------

    private IReadOnlyList<PdfOutlineNode> MapBookmarks(
        PdfDocumentProcessor processor, IList<PdfBookmark> bookmarks, int level)
    {
        return bookmarks
            .Select(bookmark => new PdfOutlineNode(
                bookmark.Title ?? string.Empty,
                level,
                ResolvePageNumber(processor, bookmark.Destination),
                MapBookmarks(processor, bookmark.Children, level + 1)))
            .ToArray();
    }

    private static int? ResolvePageNumber(PdfDocumentProcessor processor, PdfDestination? destination)
    {
        if (destination?.Page is null)
        {
            return null;
        }

        var index = processor.Document.Pages.IndexOf(destination.Page);
        return index < 0 ? null : index + 1;
    }

    private static Dictionary<int, List<PdfWordBox>> CollectWords(
        PdfDocumentProcessor processor, IReadOnlyList<int> wanted, bool includeFontInfo)
        => CollectWords(processor, wanted, includeFontInfo, int.MaxValue, out _);

    /// <summary>
    /// Walks the document's word cursor once and buckets the words by page.
    ///
    /// NextWord() is a forward-only cursor over the whole document — there is no per-page
    /// overload — so pages outside <paramref name="wanted"/> are walked and dropped rather
    /// than skipped. That is still one pass, and it is the only positioned-text API
    /// DevExpress exposes.
    /// </summary>
    private static Dictionary<int, List<PdfWordBox>> CollectWords(
        PdfDocumentProcessor processor,
        IReadOnlyList<int> wanted,
        bool includeFontInfo,
        int maxWords,
        out bool truncated)
    {
        var wantedSet = wanted.ToHashSet();
        var byPage = new Dictionary<int, List<PdfWordBox>>();
        var total = 0;
        truncated = false;

        PdfPageWord? word;
        while ((word = processor.NextWord()) is not null)
        {
            if (!wantedSet.Contains(word.PageNumber) || word.Rectangles.Count == 0)
            {
                continue;
            }

            if (total >= maxWords)
            {
                truncated = true;
                break;
            }

            var page = processor.Document.Pages[word.PageNumber - 1];
            var rectangle = word.Rectangles[0];
            var character = word.Characters.Count > 0 ? word.Characters[0] : null;

            if (!byPage.TryGetValue(word.PageNumber, out var list))
            {
                list = [];
                byPage[word.PageNumber] = list;
            }

            list.Add(new PdfWordBox(
                word.Text,
                Round(rectangle.Left),
                Round(FlipY(rectangle.Top, page)),
                Round(rectangle.Width),
                Round(rectangle.Height),
                includeFontInfo ? Round(character?.FontSize ?? 0) : null,
                includeFontInfo ? character?.Font?.FontName : null));

            total++;
        }

        return byPage;
    }

    /// <summary>
    /// PDF measures Y upward from the bottom of the page; every consumer of these tools wants
    /// reading order, so Y is reported as distance from the top edge instead.
    /// </summary>
    private static double FlipY(double topInPdfSpace, PdfPage page)
        => page.CropBox.Top - topInPdfSpace;

    private static double Round(double value) => Math.Round(value, 2);

    private static string? NullIfBlank(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

    private static int CountBookmarks(IList<PdfBookmark> bookmarks)
        => bookmarks.Count + bookmarks.Sum(b => CountBookmarks(b.Children));

    private static DXImageFormat ResolveImageFormat(string? format, string outputPath)
    {
        var name = format;
        if (string.IsNullOrWhiteSpace(name))
        {
            name = Path.GetExtension(outputPath).TrimStart('.');
        }

        return name.Trim().ToLowerInvariant() switch
        {
            "png" => DXImageFormat.Png,
            "jpg" or "jpeg" => DXImageFormat.Jpeg,
            "bmp" => DXImageFormat.Bmp,
            "gif" => DXImageFormat.Gif,
            "tif" or "tiff" => DXImageFormat.Tiff,
            _ => throw ToolError.UnsupportedImageFormat(name),
        };
    }

    /// <summary>
    /// Loads the document, runs <paramref name="read"/>, and maps DevExpress failures onto the
    /// tool error codes. An encrypted document surfaces as password_required rather than a
    /// generic parse error, because the caller can act on that one.
    /// </summary>
    private static T Read<T>(string path, Func<PdfDocumentProcessor, T> read)
    {
        try
        {
            using var processor = new PdfDocumentProcessor();
            processor.LoadDocument(path);
            return read(processor);
        }
        catch (PdfIncorrectPasswordException)
        {
            throw ToolError.PasswordRequired(path);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }
}
