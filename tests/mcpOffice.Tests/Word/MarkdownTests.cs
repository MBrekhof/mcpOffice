using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class MarkdownTests
{
    [Fact]
    public void ReadAsMarkdown_returns_headings_and_paragraphs()
    {
        var path = TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Introduction", "Heading 1");
            TestWordDocuments.AppendParagraph(document, "Background", "Heading 2");
            TestWordDocuments.AppendParagraph(document, "Plain paragraph with useful text.");
        });

        var markdown = new WordDocumentService().ReadAsMarkdown(path);

        Assert.Contains("# Introduction", markdown);
        Assert.Contains("## Background", markdown);
        Assert.Contains("Plain paragraph with useful text.", markdown);
    }
}
