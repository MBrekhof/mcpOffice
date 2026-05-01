using McpOffice.Services.Word;
using ModelContextProtocol;

namespace McpOffice.Tests.Word;

public class InsertParagraphTests
{
    [Fact]
    public void InsertParagraph_at_zero_with_heading_style_grows_outline()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-insert-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(path, "Existing body text.", overwrite: false);

            service.InsertParagraph(path, atIndex: 0, text: "Title", style: "Heading 1");

            var outline = service.GetOutline(path);
            Assert.Single(outline);
            Assert.Equal(1, outline[0].Level);
            Assert.Equal("Title", outline[0].Text);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void InsertParagraph_throws_index_out_of_range_when_index_is_too_large()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-insert-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateBlank(path, overwrite: false);

            Action act = () => service.InsertParagraph(path, atIndex: 999, text: "X", style: null);
            var ex = Assert.Throws<McpException>(act);
            Assert.Contains("index_out_of_range", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
