using DevExpress.XtraRichEdit.API.Native;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class CommentsTests
{
    [Fact]
    public void ListComments_returns_authors_text_and_anchor_text()
    {
        var path = TestWordDocuments.Create(document =>
        {
            AppendComment(document, "First anchor", "Alice", "First note");
            AppendComment(document, "Second anchor", "Bob", "Second note");
        });

        var comments = new WordDocumentService().ListComments(path);

        Assert.Equal(2, comments.Count);
        Assert.Contains(comments, c => c.Author == "Alice" && c.Text == "First note" && c.AnchorText == "First anchor");
        Assert.Contains(comments, c => c.Author == "Bob" && c.Text == "Second note" && c.AnchorText == "Second anchor");
    }

    private static void AppendComment(Document document, string anchorText, string author, string commentText)
    {
        var anchor = document.AppendText(anchorText);
        document.AppendText(Environment.NewLine);
        var comment = document.Comments.Create(anchor, author);
        var body = comment.BeginUpdate();
        body.AppendText(commentText);
        comment.EndUpdate(body);
    }
}
