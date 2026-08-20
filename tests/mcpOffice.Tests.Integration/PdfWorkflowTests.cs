using ModelContextProtocol.Protocol;

namespace McpOffice.Tests.Integration;

public class PdfWorkflowTests
{
    [Fact]
    public async Task Convert_then_read_text_via_stdio()
    {
        var docx = TempPath(".docx");
        var pdf = TempPath(".pdf");
        try
        {
            await using var harness = await ServerHarness.StartAsync();

            await CallTextAsync(harness, "word_create_from_markdown", new Dictionary<string, object?>
            {
                ["path"] = docx,
                ["markdown"] = "# Overzichtsrapport\n\nChloride 120 mg/l",
                ["overwrite"] = false
            });

            await CallTextAsync(harness, "word_convert", new Dictionary<string, object?>
            {
                ["inputPath"] = docx,
                ["outputPath"] = pdf
            });

            var textJson = await CallTextAsync(harness, "pdf_read_text", new Dictionary<string, object?>
            {
                ["path"] = pdf
            });

            Assert.Contains("Overzichtsrapport", textJson);
            Assert.Contains("Chloride", textJson);
        }
        finally
        {
            DeleteIfExists(docx);
            DeleteIfExists(pdf);
        }
    }

    [Fact]
    public async Task Render_page_to_png_via_stdio()
    {
        var docx = TempPath(".docx");
        var pdf = TempPath(".pdf");
        var png = TempPath(".png");
        try
        {
            await using var harness = await ServerHarness.StartAsync();

            await CallTextAsync(harness, "word_create_from_markdown", new Dictionary<string, object?>
            {
                ["path"] = docx,
                ["markdown"] = "# Render me",
                ["overwrite"] = false
            });

            await CallTextAsync(harness, "word_convert", new Dictionary<string, object?>
            {
                ["inputPath"] = docx,
                ["outputPath"] = pdf
            });

            var renderJson = await CallTextAsync(harness, "pdf_render_page", new Dictionary<string, object?>
            {
                ["path"] = pdf,
                ["pageNumber"] = 1,
                ["outputPath"] = png,
                ["dpi"] = 96
            });

            Assert.Contains("pixelWidth", renderJson, StringComparison.OrdinalIgnoreCase);
            Assert.True(File.Exists(png));
            // PNG magic number - proves a real image was written, not just a path echoed back.
            Assert.Equal([0x89, 0x50, 0x4E, 0x47], File.ReadAllBytes(png)[..4]);
        }
        finally
        {
            DeleteIfExists(docx);
            DeleteIfExists(pdf);
            DeleteIfExists(png);
        }
    }

    private static async Task<string> CallTextAsync(
        ServerHarness harness,
        string toolName,
        IReadOnlyDictionary<string, object?> arguments)
    {
        var result = await harness.Client.CallToolAsync(toolName, arguments);
        return result.Content.OfType<TextContentBlock>().Single().Text;
    }

    private static string TempPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"mcpoffice-integration-{Guid.NewGuid():N}{extension}");

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
