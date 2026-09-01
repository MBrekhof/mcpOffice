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
    public void MailMerge_throws_file_exists_when_output_exists_and_overwrite_not_requested()
    {
        var template = Path.Combine(Path.GetTempPath(), $"mcpoffice-tpl-{Guid.NewGuid():N}.docx");
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-out-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(template, "Hello {{firstName}}!", overwrite: false);
            service.MailMerge(template, output, "{\"firstName\":\"Ada\"}");

            Action act = () => service.MailMerge(template, output, "{\"firstName\":\"Bob\"}");
            var ex = Assert.Throws<McpException>(act);
            Assert.Contains("file_exists", ex.Message);
        }
        finally
        {
            if (File.Exists(template)) File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void MailMerge_with_overwrite_replaces_an_existing_output()
    {
        // WORD-001: regenerating into the same path (CSV price list -> merge -> pdf, every run).
        var template = Path.Combine(Path.GetTempPath(), $"mcpoffice-tpl-{Guid.NewGuid():N}.docx");
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-out-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateFromMarkdown(template, "Hello {{firstName}}!", overwrite: false);
            service.MailMerge(template, output, "{\"firstName\":\"Ada\"}");

            service.MailMerge(template, output, "{\"firstName\":\"Bob\"}", overwrite: true);

            Assert.Contains("Hello Bob!", service.ReadAsMarkdown(output));
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
