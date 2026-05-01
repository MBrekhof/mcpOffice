using McpOffice.Models;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class CreateFromMarkdownTests
{
    [Fact]
    public void CreateFromMarkdown_writes_heading_and_bold_run()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-md-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(path, "# Title\n\nHello **world**", overwrite: false);

            var structured = service.ReadStructured(path);
            Assert.Contains(structured.Blocks, b => b is HeadingBlock h && h.Text == "Title");

            var paragraph = structured.Blocks.OfType<ParagraphBlock>().Single();
            Assert.Contains(paragraph.Runs, r => r.Bold && r.Text.Contains("world"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
