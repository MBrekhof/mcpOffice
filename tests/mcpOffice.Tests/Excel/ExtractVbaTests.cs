using McpOffice.Services.Excel;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel;

public class ExtractVbaTests
{
    [Fact]
    public void ExtractVba_throws_file_not_found_for_missing_workbook()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"mcpoffice-missing-{Guid.NewGuid():N}.xlsm");

        var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().ExtractVba(missing));

        Assert.Contains("file_not_found", ex.Message);
    }

    [Fact]
    public void ExtractVba_returns_HasVbaProject_false_for_xlsx_without_macros()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Name = "Data";
            workbook.Worksheets[0].Cells["A1"].Value = "x";
        });

        try
        {
            var project = new ExcelWorkbookService().ExtractVba(path);

            Assert.False(project.HasVbaProject);
            Assert.Empty(project.Modules);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
