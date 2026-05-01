using McpOffice.Models;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class CreateFromMarkdownTests
{
    [Fact]
    public void CreateFromMarkdown_writes_heading_bold_and_italic_runs()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-md-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(path, "# Title\n\nHello **world** and *reader*", overwrite: false);

            var structured = service.ReadStructured(path);
            Assert.Contains(structured.Blocks, b => b is HeadingBlock h && h.Text == "Title");

            var paragraph = structured.Blocks.OfType<ParagraphBlock>().Single();
            Assert.Contains(paragraph.Runs, r => r.Bold && r.Text.Contains("world"));
            Assert.Contains(paragraph.Runs, r => r.Italic && r.Text.Contains("reader"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void CreateFromMarkdown_writes_tables_lists_and_code_blocks()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-md-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(
                path,
                """
                # Rich Markdown

                - First bullet
                - Second bullet

                | Name | Value |
                | ---- | ----- |
                | Alpha | 1 |
                | Beta | 2 |

                ```csharp
                Console.WriteLine("hello");
                ```
                """,
                overwrite: false);

            var structured = service.ReadStructured(path);
            Assert.Contains(structured.Blocks, b => b is HeadingBlock h && h.Text == "Rich Markdown");
            Assert.Contains(structured.Blocks.OfType<ParagraphBlock>(), p => p.Runs.Any(r => r.Text.Contains("First bullet")));
            Assert.Contains(structured.Blocks.OfType<ParagraphBlock>(), p => p.Runs.Any(r => r.Text.Contains("Console.WriteLine")));

            var table = Assert.Single(structured.Tables);
            Assert.Equal("Name", table.Rows[0][0]);
            Assert.Equal("Beta", table.Rows[2][0]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
