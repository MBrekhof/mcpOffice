using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class MetadataTests
{
    [Fact]
    public void Metadata_returns_core_properties_and_counts()
    {
        var created = new DateTime(2026, 4, 30, 10, 0, 0, DateTimeKind.Utc);
        var modified = new DateTime(2026, 4, 30, 11, 0, 0, DateTimeKind.Utc);
        var path = TestWordDocuments.Create(document =>
        {
            document.DocumentProperties.Author = "Martin";
            document.DocumentProperties.Title = "Metadata Fixture";
            document.DocumentProperties.Subject = "MCP Office";
            document.DocumentProperties.Keywords = "mcp,office,word";
            document.DocumentProperties.Created = created;
            document.DocumentProperties.Modified = modified;
            document.DocumentProperties.Revision = 7;

            TestWordDocuments.AppendParagraph(document, "Introduction", "Heading 1");
            TestWordDocuments.AppendParagraph(document, "Hello world from Word.");
        });

        var metadata = new WordDocumentService().GetMetadata(path);

        Assert.Equal("Martin", metadata.Author);
        Assert.Equal("Metadata Fixture", metadata.Title);
        Assert.Equal("MCP Office", metadata.Subject);
        Assert.Equal("mcp,office,word", metadata.Keywords);
        Assert.Equal(7, metadata.RevisionCount);
        Assert.True(metadata.PageCount >= 1);
        Assert.True(metadata.WordCount >= 5);
    }
}
