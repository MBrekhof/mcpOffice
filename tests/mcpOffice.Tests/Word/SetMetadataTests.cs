using McpOffice.Services.Word;
using ModelContextProtocol;

namespace McpOffice.Tests.Word;

public class SetMetadataTests
{
    [Fact]
    public void SetMetadata_writes_author_title_subject_keywords()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-meta-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateBlank(path, overwrite: false);

            service.SetMetadata(path, new Dictionary<string, string>
            {
                ["author"] = "Bob",
                ["title"] = "My Doc",
                ["subject"] = "Testing",
                ["keywords"] = "alpha,beta"
            });

            var meta = service.GetMetadata(path);
            Assert.Equal("Bob", meta.Author);
            Assert.Equal("My Doc", meta.Title);
            Assert.Equal("Testing", meta.Subject);
            Assert.Equal("alpha,beta", meta.Keywords);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SetMetadata_throws_unsupported_format_for_unknown_key()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-meta-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateBlank(path, overwrite: false);

            Action act = () => service.SetMetadata(path, new Dictionary<string, string>
            {
                ["color"] = "blue"
            });
            var ex = Assert.Throws<McpException>(act);
            Assert.Contains("unsupported_format", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
