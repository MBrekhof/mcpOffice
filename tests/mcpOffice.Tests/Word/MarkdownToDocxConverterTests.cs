using DevExpress.XtraRichEdit;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class MarkdownToDocxConverterTests
{
    [Fact]
    public void Apply_with_empty_markdown_does_not_throw()
    {
        using var server = new RichEditDocumentServer();
        // Should not throw
        MarkdownToDocxConverter.Apply(server.Document, string.Empty, null);
    }

    [Fact]
    public void Empty_markdown_produces_empty_document()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "", null);
        Assert.Single(server.Document.Paragraphs);
        Assert.Equal(string.Empty, server.Document.GetText(server.Document.Range).Trim());
    }

    [Fact]
    public void Plain_paragraphs_become_paragraphs_with_literal_text()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "hello world\n\nsecond para", null);
        var text = server.Document.GetText(server.Document.Range).Trim();
        Assert.Contains("hello world", text);
        Assert.Contains("second para", text);
        Assert.True(server.Document.Paragraphs.Count >= 2,
            $"expected ≥2 paragraphs, got {server.Document.Paragraphs.Count}");
    }
}
