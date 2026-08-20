using McpOffice.Services.Pdf;
using ModelContextProtocol;

namespace McpOffice.Tests.Pdf;

public class PdfDocumentServiceTests
{
    private static readonly PdfDocumentService Service = new();

    [Fact]
    public void GetMetadata_reports_pages_and_geometry()
    {
        var pdf = TestPdfDocuments.Create("Hello metadata");
        try
        {
            var meta = Service.GetMetadata(pdf);

            Assert.Equal(1, meta.PageCount);
            var page = Assert.Single(meta.Pages);
            Assert.Equal(1, page.PageNumber);
            Assert.True(page.Width > 0);
            Assert.True(page.Height > 0);
            Assert.False(string.IsNullOrWhiteSpace(meta.Version));
            Assert.Equal(0, meta.BookmarkCount);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadText_returns_the_text_of_each_page()
    {
        var pdf = TestPdfDocuments.Create("Koloniegetal bij 22 C", "Chloride 120 mg/l");
        try
        {
            var result = Service.ReadText(pdf, pageRange: null, preserveLayout: false, maxChars: 200_000);

            Assert.Equal(1, result.PageCount);
            Assert.False(result.Truncated);
            var text = Assert.Single(result.Pages).Text;
            Assert.Contains("Koloniegetal", text);
            Assert.Contains("Chloride", text);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadText_honours_page_range()
    {
        var pdf = TestPdfDocuments.CreateMultiPage(3);
        try
        {
            var result = Service.ReadText(pdf, pageRange: "2", preserveLayout: false, maxChars: 200_000);

            Assert.Equal(3, result.PageCount);
            var page = Assert.Single(result.Pages);
            Assert.Equal(2, page.PageNumber);
            Assert.Contains("Page marker 2", page.Text);
            Assert.DoesNotContain("Page marker 1", page.Text);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadText_truncates_at_maxChars()
    {
        var pdf = TestPdfDocuments.Create("A reasonably long line of text that will be cut short");
        try
        {
            var result = Service.ReadText(pdf, pageRange: null, preserveLayout: false, maxChars: 10);

            Assert.True(result.Truncated);
            Assert.Equal(10, result.CharCount);
            Assert.Equal(10, Assert.Single(result.Pages).Text.Length);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadText_with_preserveLayout_keeps_the_words()
    {
        var pdf = TestPdfDocuments.Create("Parameter Eenheid Norm");
        try
        {
            var result = Service.ReadText(pdf, pageRange: null, preserveLayout: true, maxChars: 200_000);

            var text = Assert.Single(result.Pages).Text;
            Assert.Contains("Parameter", text);
            Assert.Contains("Eenheid", text);
            Assert.Contains("Norm", text);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadLayout_line_mode_returns_positioned_lines()
    {
        var pdf = TestPdfDocuments.Create("First line here", "Second line here");
        try
        {
            var result = Service.ReadLayout(pdf, null, "line", includeFontInfo: false, maxWords: 50_000);

            Assert.Equal("line", result.Granularity);
            var page = Assert.Single(result.Pages);
            Assert.NotNull(page.Lines);
            Assert.Null(page.Words);
            Assert.True(page.Lines!.Count >= 2);

            // Reading order: y grows downward, so line 1 sits above line 2.
            Assert.True(page.Lines[0].Y < page.Lines[1].Y);
            Assert.Contains("First", page.Lines[0].Text);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadLayout_word_mode_returns_words_with_boxes()
    {
        var pdf = TestPdfDocuments.Create("Chloride 120");
        try
        {
            var result = Service.ReadLayout(pdf, null, "word", includeFontInfo: true, maxWords: 50_000);

            var page = Assert.Single(result.Pages);
            Assert.NotNull(page.Words);
            Assert.Null(page.Lines);

            var word = Assert.Single(page.Words!, w => w.Text == "Chloride");
            Assert.True(word.Width > 0);
            Assert.True(word.Height > 0);
            Assert.True(word.Y >= 0 && word.Y <= page.Height);
            Assert.NotNull(word.FontSize);
            Assert.False(string.IsNullOrWhiteSpace(word.FontName));
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadLayout_omits_font_info_by_default()
    {
        var pdf = TestPdfDocuments.Create("Chloride");
        try
        {
            var result = Service.ReadLayout(pdf, null, "word", includeFontInfo: false, maxWords: 50_000);

            var word = Assert.Single(Assert.Single(result.Pages).Words!);
            Assert.Null(word.FontSize);
            Assert.Null(word.FontName);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadLayout_rejects_an_unknown_granularity()
    {
        var pdf = TestPdfDocuments.Create("x");
        try
        {
            var ex = Assert.Throws<McpException>(
                () => Service.ReadLayout(pdf, null, "paragraph", false, 50_000));
            Assert.StartsWith("[unsupported_format]", ex.Message);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ReadLayout_truncates_at_maxWords()
    {
        var pdf = TestPdfDocuments.Create("one two three four five six seven eight");
        try
        {
            var result = Service.ReadLayout(pdf, null, "word", false, maxWords: 3);

            Assert.True(result.Truncated);
            Assert.Equal(3, result.WordCount);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void FindText_returns_hits_with_boxes()
    {
        var pdf = TestPdfDocuments.Create("Monsterpunt: Avebe TAK", "Monsterpunt: Ruwwater");
        try
        {
            var result = Service.FindText(pdf, "Monsterpunt", caseSensitive: false, wholeWords: false, maxResults: 500);

            Assert.True(result.HitCount > 0);
            Assert.False(result.Truncated);
            Assert.All(result.Hits, hit =>
            {
                Assert.Equal(1, hit.PageNumber);
                Assert.True(hit.Width > 0);
                Assert.True(hit.Height > 0);
            });
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void FindText_returns_nothing_for_an_absent_term()
    {
        var pdf = TestPdfDocuments.Create("Chloride");
        try
        {
            var result = Service.FindText(pdf, "Sulfaat", false, false, 500);

            Assert.Equal(0, result.HitCount);
            Assert.Empty(result.Hits);
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void RenderPage_writes_an_image()
    {
        var pdf = TestPdfDocuments.Create("Render me");
        var png = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.png");
        try
        {
            var result = Service.RenderPage(pdf, 1, png, dpi: 96, format: null, overwrite: false);

            Assert.True(File.Exists(png));
            Assert.Equal(png, result.OutputPath);
            Assert.Equal(1, result.PageNumber);
            Assert.True(result.PixelWidth > 0);
            Assert.True(result.PixelHeight > 0);
            Assert.True(result.BytesWritten > 0);
        }
        finally { TestPdfDocuments.Delete(pdf, png); }
    }

    [Fact]
    public void RenderPage_rejects_a_page_that_does_not_exist()
    {
        var pdf = TestPdfDocuments.Create("Only one page");
        var png = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.png");
        try
        {
            var ex = Assert.Throws<McpException>(() => Service.RenderPage(pdf, 9, png, 96, null, false));
            Assert.StartsWith("[page_not_found]", ex.Message);
        }
        finally { TestPdfDocuments.Delete(pdf, png); }
    }

    [Fact]
    public void RenderPage_rejects_an_unknown_image_format()
    {
        var pdf = TestPdfDocuments.Create("x");
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.xyz");
        try
        {
            var ex = Assert.Throws<McpException>(() => Service.RenderPage(pdf, 1, output, 96, null, false));
            Assert.StartsWith("[unsupported_format]", ex.Message);
        }
        finally { TestPdfDocuments.Delete(pdf, output); }
    }

    [Fact]
    public void RenderPage_rejects_an_out_of_band_dpi()
    {
        var pdf = TestPdfDocuments.Create("x");
        var png = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.png");
        try
        {
            var ex = Assert.Throws<McpException>(() => Service.RenderPage(pdf, 1, png, 5000, null, false));
            Assert.StartsWith("[invalid_render_option]", ex.Message);
        }
        finally { TestPdfDocuments.Delete(pdf, png); }
    }

    [Fact]
    public void RenderPage_will_not_clobber_an_existing_file()
    {
        var pdf = TestPdfDocuments.Create("x");
        var png = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.png");
        File.WriteAllText(png, "already here");
        try
        {
            var ex = Assert.Throws<McpException>(() => Service.RenderPage(pdf, 1, png, 96, null, overwrite: false));
            Assert.StartsWith("[file_exists]", ex.Message);
        }
        finally { TestPdfDocuments.Delete(pdf, png); }
    }

    [Fact]
    public void GetOutline_is_empty_when_there_are_no_bookmarks()
    {
        var pdf = TestPdfDocuments.Create("No bookmarks here");
        try
        {
            Assert.Empty(Service.GetOutline(pdf));
        }
        finally { TestPdfDocuments.Delete(pdf); }
    }

    [Fact]
    public void ExtractImages_returns_nothing_for_a_text_only_pdf()
    {
        var pdf = TestPdfDocuments.Create("Text only");
        var dir = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}");
        try
        {
            var result = Service.ExtractImages(pdf, dir, null, minPixelSize: 16, maxImages: 200, overwrite: false);

            Assert.Equal(0, result.ImageCount);
            Assert.Empty(result.Images);
            Assert.Equal(dir, result.OutputDirectory);
        }
        finally
        {
            TestPdfDocuments.Delete(pdf);
            if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Missing_file_is_file_not_found()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.pdf");

        var ex = Assert.Throws<McpException>(() => Service.GetMetadata(missing));
        Assert.StartsWith("[file_not_found]", ex.Message);
    }

    [Fact]
    public void Relative_path_is_invalid_path()
    {
        var ex = Assert.Throws<McpException>(() => Service.GetMetadata("relative.pdf"));
        Assert.StartsWith("[invalid_path]", ex.Message);
    }

    [Fact]
    public void A_non_pdf_file_is_a_parse_error()
    {
        var notPdf = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.pdf");
        File.WriteAllText(notPdf, "this is not a pdf");
        try
        {
            var ex = Assert.Throws<McpException>(() => Service.GetMetadata(notPdf));
            Assert.StartsWith("[parse_error]", ex.Message);
        }
        finally { TestPdfDocuments.Delete(notPdf); }
    }
}
