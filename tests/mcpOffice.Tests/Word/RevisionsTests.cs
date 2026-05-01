using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class RevisionsTests
{
    [Fact]
    public void ListRevisions_returns_inserted_text_with_author_and_insert_type()
    {
        var path = TestWordDocuments.Create(server =>
        {
            var document = server.Document;
            document.AppendText("Original. ");

            server.Options.Annotations.Author = "Reviewer";
            document.TrackChanges.Enabled = true;
            document.AppendText("Inserted text.");
        });

        var revisions = new WordDocumentService().ListRevisions(path);

        Assert.Contains(revisions, r =>
            r.Type == "insert" &&
            r.Author == "Reviewer" &&
            r.Text.Contains("Inserted text."));
    }
}
