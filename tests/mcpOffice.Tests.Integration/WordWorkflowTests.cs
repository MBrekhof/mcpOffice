using ModelContextProtocol.Protocol;
using System.Text;

namespace McpOffice.Tests.Integration;

public class WordWorkflowTests
{
    [Fact]
    public async Task Read_markdown_round_trip_via_stdio()
    {
        var path = TempPath(".docx");
        try
        {
            await using var harness = await ServerHarness.StartAsync();
            await CallTextAsync(
                harness,
                "word_create_from_markdown",
                new Dictionary<string, object?>
                {
                    ["path"] = path,
                    ["markdown"] = "# Round Trip\n\nHello from stdio.",
                    ["overwrite"] = false
                });

            var markdown = await CallTextAsync(
                harness,
                "word_read_markdown",
                new Dictionary<string, object?> { ["path"] = path });

            Assert.Contains("# Round Trip", markdown);
            Assert.Contains("Hello from stdio.", markdown);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public async Task Create_then_outline_via_stdio()
    {
        var path = TempPath(".docx");
        try
        {
            await using var harness = await ServerHarness.StartAsync();
            var createdPath = await CallTextAsync(
                harness,
                "word_create_from_markdown",
                new Dictionary<string, object?>
                {
                    ["path"] = path,
                    ["markdown"] = "# Stdio Title\n\nBody text.",
                    ["overwrite"] = false
                });

            Assert.Equal(path, createdPath);

            var outlineJson = await CallTextAsync(
                harness,
                "word_get_outline",
                new Dictionary<string, object?> { ["path"] = path });

            Assert.Contains("Stdio Title", outlineJson);
            Assert.Contains("\"level\":1", outlineJson, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public async Task Convert_to_pdf_via_stdio()
    {
        var inputPath = TempPath(".docx");
        var outputPath = TempPath(".pdf");
        try
        {
            await using var harness = await ServerHarness.StartAsync();
            await CallTextAsync(
                harness,
                "word_create_from_markdown",
                new Dictionary<string, object?>
                {
                    ["path"] = inputPath,
                    ["markdown"] = "# PDF Title\n\nConverted through MCP.",
                    ["overwrite"] = false
                });

            var convertedPath = await CallTextAsync(
                harness,
                "word_convert",
                new Dictionary<string, object?>
                {
                    ["inputPath"] = inputPath,
                    ["outputPath"] = outputPath
                });

            Assert.Equal(outputPath, convertedPath);
            Assert.True(File.Exists(outputPath));
            Assert.Equal("%PDF-", Encoding.ASCII.GetString(File.ReadAllBytes(outputPath)[..5]));
        }
        finally
        {
            DeleteIfExists(inputPath);
            DeleteIfExists(outputPath);
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
