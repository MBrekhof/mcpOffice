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

    [Fact]
    public void Unordered_list_produces_bulleted_paragraphs()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "- a\n- b\n- c", null);
        var listParas = server.Document.Paragraphs
            .Where(p => p.ListIndex >= 0)
            .ToList();
        Assert.Equal(3, listParas.Count);
        Assert.All(listParas, p => Assert.Equal(0, p.ListLevel));
    }

    [Fact]
    public void Ordered_list_produces_numbered_paragraphs()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "1. one\n2. two", null);
        var listParas = server.Document.Paragraphs
            .Where(p => p.ListIndex >= 0)
            .ToList();
        Assert.Equal(2, listParas.Count);
    }

    [Fact]
    public void Nested_list_indents_per_depth()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "- outer\n  - inner", null);
        var paras = server.Document.Paragraphs.Where(p => p.ListIndex >= 0).ToList();
        Assert.Equal(2, paras.Count);
        Assert.Equal(0, paras[0].ListLevel);
        Assert.Equal(1, paras[1].ListLevel);
    }
}
