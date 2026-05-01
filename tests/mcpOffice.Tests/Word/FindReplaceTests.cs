using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class FindReplaceTests
{
    [Fact]
    public void FindReplace_replaces_all_occurrences_and_returns_count()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-fr-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(path, "hello hello", overwrite: false);

            var result = service.FindReplace(path, "hello", "hi", useRegex: false, matchCase: false);

            Assert.Equal(2, result.Replacements);
            var markdown = service.ReadAsMarkdown(path);
            Assert.Contains("hi hi", markdown);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
