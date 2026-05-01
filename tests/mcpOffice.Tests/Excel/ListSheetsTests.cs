using McpOffice.Services.Excel;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel;

public class ListSheetsTests
{
    [Fact]
    public void ListSheets_returns_sheet_names_visibility_and_used_ranges()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Name = "Data";
            workbook.Worksheets[0].Cells["A1"].Value = "Name";
            workbook.Worksheets[0].Cells["B2"].Value = 42;

            var hidden = workbook.Worksheets.Add("Hidden");
            hidden.Visible = false;
            hidden.Cells["C3"].Value = "x";
        });

        try
        {
            var sheets = new ExcelWorkbookService().ListSheets(path);

            Assert.Equal(2, sheets.Count);
            Assert.Equal("Data", sheets[0].Name);
            Assert.True(sheets[0].Visible);
            Assert.Equal("A1:B2", sheets[0].UsedRange);
            Assert.Equal(2, sheets[0].RowCount);
            Assert.Equal(2, sheets[0].ColumnCount);

            Assert.Equal("Hidden", sheets[1].Name);
            Assert.False(sheets[1].Visible);
            Assert.Equal("C3", sheets[1].UsedRange);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ListSheets_throws_file_not_found_for_missing_workbook()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"mcpoffice-missing-{Guid.NewGuid():N}.xlsx");

        var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().ListSheets(missing));

        Assert.Contains("file_not_found", ex.Message);
    }
}
