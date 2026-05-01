using DevExpress.XtraRichEdit;
using McpOffice.Services.Word;
using ModelContextProtocol;
using RichEditFormat = DevExpress.XtraRichEdit.DocumentFormat;

namespace McpOffice.Tests.Word;

public class CreateBlankTests
{
    [Fact]
    public void CreateBlank_writes_a_loadable_docx_and_returns_path()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-blank-{Guid.NewGuid():N}.docx");
        try
        {
            var returned = new WordDocumentService().CreateBlank(path, overwrite: false);

            Assert.Equal(path, returned);
            Assert.True(File.Exists(path));

            using var server = new RichEditDocumentServer();
            server.LoadDocument(path, RichEditFormat.OpenXml);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void CreateBlank_throws_file_exists_without_overwrite()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-blank-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateBlank(path, overwrite: false);

            Action act = () => service.CreateBlank(path, overwrite: false);
            var ex = Assert.Throws<McpException>(act);
            Assert.Contains("file_exists", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void CreateBlank_overwrites_when_overwrite_true()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-blank-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateBlank(path, overwrite: false);
            service.CreateBlank(path, overwrite: true);

            Assert.True(File.Exists(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
