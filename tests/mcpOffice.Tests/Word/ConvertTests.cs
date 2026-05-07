using DevExpress.XtraRichEdit;
using McpOffice.Services.Word;
using ModelContextProtocol;
using System.Text;

namespace McpOffice.Tests.Word;

public class ConvertTests
{
    [Theory]
    [InlineData("pdf")]
    [InlineData("html")]
    [InlineData("rtf")]
    [InlineData("txt")]
    [InlineData("md")]
    [InlineData("docx")]
    public void Convert_writes_expected_format_from_extension(string extension)
    {
        var input = CreateInputDocument();
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-convert-{Guid.NewGuid():N}.{extension}");

        try
        {
            var result = new WordDocumentService().Convert(input, output, format: null);

            Assert.Equal(output, result);
            Assert.True(File.Exists(output));
            AssertFormatLooksRight(output, extension);
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

    [Fact]
    public void Convert_accepts_explicit_markdown_format()
    {
        var input = CreateInputDocument();
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-convert-{Guid.NewGuid():N}.out");

        try
        {
            new WordDocumentService().Convert(input, output, "markdown");

            var markdown = File.ReadAllText(output, Encoding.UTF8);
            Assert.Contains("# Convert Me", markdown);
            Assert.Contains("Plain text", markdown);
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

    [Fact]
    public void Convert_md_input_to_docx_preserves_inline_code_and_tables()
    {
        var mdPath = Path.Combine(Path.GetTempPath(), $"mcpoffice-mdinput-{Guid.NewGuid():N}.md");
        var docxPath = Path.Combine(Path.GetTempPath(), $"mcpoffice-mdout-{Guid.NewGuid():N}.docx");

        try
        {
            File.WriteAllText(mdPath,
                "# Title\n\nUse `Foo()` to call. Tables:\n\n| A | B |\n|---|---|\n| 1 | 2 |\n",
                Encoding.UTF8);

            var svc = new WordDocumentService();
            svc.Convert(mdPath, docxPath, format: null);

            Assert.True(File.Exists(docxPath));
            using var server = new DevExpress.XtraRichEdit.RichEditDocumentServer();
            server.LoadDocument(docxPath, DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            Assert.Contains("Foo()", server.Document.GetText(server.Document.Range));
            Assert.True(server.Document.Tables.Count >= 1);
        }
        finally
        {
            DeleteIfExists(mdPath);
            DeleteIfExists(docxPath);
        }
    }

    [Fact]
    public void Convert_rejects_unknown_format()
    {
        var input = CreateInputDocument();
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-convert-{Guid.NewGuid():N}.xyz");

        try
        {
            var ex = Assert.Throws<McpException>(() =>
                new WordDocumentService().Convert(input, output, "xyz"));

            Assert.Contains("unsupported_format", ex.Message);
            Assert.False(File.Exists(output));
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

    private static string CreateInputDocument() =>
        TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Convert Me", "Heading 1");
            TestWordDocuments.AppendParagraph(document, "Plain text for conversion.");
        });

    private static void AssertFormatLooksRight(string path, string extension)
    {
        switch (extension)
        {
            case "pdf":
                Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(File.ReadAllBytes(path)[..5]));
                break;
            case "html":
                Assert.Contains("<html", File.ReadAllText(path).ToLowerInvariant());
                break;
            case "rtf":
                Assert.StartsWith(@"{\rtf", File.ReadAllText(path));
                break;
            case "txt":
                Assert.Contains("Convert Me", File.ReadAllText(path));
                break;
            case "md":
                Assert.Contains("# Convert Me", File.ReadAllText(path, Encoding.UTF8));
                break;
            case "docx":
                Assert.Equal([0x50, 0x4B, 0x03, 0x04], File.ReadAllBytes(path)[..4]);
                break;
        }
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
