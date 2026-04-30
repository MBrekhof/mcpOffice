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
        return TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Introduction", "Heading 1");
            TestWordDocuments.AppendParagraph(document, "Background", "Heading 2");
            TestWordDocuments.AppendParagraph(document, "Plain paragraph");
            TestWordDocuments.AppendParagraph(document, "Conclusion", "Heading 1");
        });
    }
}
