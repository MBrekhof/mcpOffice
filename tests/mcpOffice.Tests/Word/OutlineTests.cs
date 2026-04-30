using DevExpress.XtraRichEdit;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class OutlineTests
{
    [Fact]
    public void Outline_returns_heading_tree()
    {
        var path = CreateHeadingsOnlyDocument();
        var service = new WordDocumentService();

        var nodes = service.GetOutline(path);

        Assert.Collection(
            nodes,
            node =>
            {
                Assert.Equal(1, node.Level);
                Assert.Equal("Introduction", node.Text);
            },
            node =>
            {
                Assert.Equal(2, node.Level);
                Assert.Equal("Background", node.Text);
            },
            node =>
            {
                Assert.Equal(1, node.Level);
                Assert.Equal("Conclusion", node.Text);
            });
    }

    private static string CreateHeadingsOnlyDocument()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-headings-{Guid.NewGuid():N}.docx");

        using var server = new RichEditDocumentServer();
        var document = server.Document;
        EnsureParagraphStyle(document, "Heading 1");
        EnsureParagraphStyle(document, "Heading 2");

        AppendParagraph(document, "Introduction", "Heading 1");
        AppendParagraph(document, "Background", "Heading 2");
        AppendParagraph(document, "Plain paragraph", null);
        AppendParagraph(document, "Conclusion", "Heading 1");

        server.SaveDocument(path, DocumentFormat.OpenXml);
        return path;
    }

    private static void EnsureParagraphStyle(DevExpress.XtraRichEdit.API.Native.Document document, string styleName)
    {
        if (document.ParagraphStyles[styleName] is not null)
        {
            return;
        }

        var style = document.ParagraphStyles.CreateNew();
        style.Name = styleName;
        document.ParagraphStyles.Add(style);
    }

    private static void AppendParagraph(DevExpress.XtraRichEdit.API.Native.Document document, string text, string? styleName)
    {
        var range = document.AppendText(text + Environment.NewLine);
        var paragraph = document.Paragraphs.Get(range).First();
        if (styleName is not null)
        {
            paragraph.Style = document.ParagraphStyles[styleName];
        }
    }
}
