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

    [Fact]
    public void Headings_1_through_6_get_correct_paragraph_style()
    {
        var md = "# h1\n\n## h2\n\n### h3\n\n#### h4\n\n##### h5\n\n###### h6";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        var headingParas = server.Document.Paragraphs
            .Where(p => p.Style?.Name?.StartsWith("Heading ") == true)
            .ToList();

        Assert.Equal(6, headingParas.Count);
        for (int i = 0; i < 6; i++)
        {
            Assert.Equal($"Heading {i + 1}", headingParas[i].Style!.Name);
        }
    }
}
