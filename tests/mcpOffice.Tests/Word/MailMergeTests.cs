using McpOffice.Services.Word;
using ModelContextProtocol;

namespace McpOffice.Tests.Word;

public class MailMergeTests
{
    [Fact]
    public void MailMerge_substitutes_tokens_into_a_new_output_doc()
    {
        var template = Path.Combine(Path.GetTempPath(), $"mcpoffice-tpl-{Guid.NewGuid():N}.docx");
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-out-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(template, "Hello {{firstName}}!", overwrite: false);

            service.MailMerge(template, output, "{\"firstName\":\"Ada\"}");

            var markdown = service.ReadAsMarkdown(output);
            Assert.Contains("Hello Ada!", markdown);
        }
        finally
        {
            if (File.Exists(template)) File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void MailMerge_throws_merge_field_missing_when_data_lacks_a_token()
    {
        var template = Path.Combine(Path.GetTempPath(), $"mcpoffice-tpl-{Guid.NewGuid():N}.docx");
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-out-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(template, "Dear {{firstName}} {{lastName}}", overwrite: false);

            Action act = () => service.MailMerge(template, output, "{\"firstName\":\"Ada\"}");
            var ex = Assert.Throws<McpException>(act);
            Assert.Contains("merge_field_missing", ex.Message);
            Assert.Contains("lastName", ex.Message);
        }
        finally
        {
            if (File.Exists(template)) File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }
}
