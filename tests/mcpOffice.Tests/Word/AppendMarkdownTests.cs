using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class AppendMarkdownTests
{
    [Fact]
    public void AppendMarkdown_adds_a_heading_to_an_existing_blank_doc()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-append-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateBlank(path, overwrite: false);
            service.AppendMarkdown(path, "# H");

            var outline = service.GetOutline(path);
            Assert.Single(outline);
            Assert.Equal(1, outline[0].Level);
            Assert.Equal("H", outline[0].Text);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
