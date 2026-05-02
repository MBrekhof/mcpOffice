using McpOffice.Services.Excel;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel;

public class ListDefinedNamesTests
{
    [Fact]
    public void ListDefinedNames_returns_workbook_and_sheet_scoped_names()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            var data = workbook.Worksheets[0];
            data.Name = "Data";
            data.Cells["A1"].Value = 10;
            data.Cells["B1"].Value = 20;

            workbook.DefinedNames.Add("TaxRate", "=0.21");
            data.DefinedNames.Add("Range1", "=Data!$A$1:$B$1");
        });

        try
        {
            var names = new ExcelWorkbookService().ListDefinedNames(path);

            Assert.Equal(2, names.Count);

            var taxRate = names.Single(n => n.Name == "TaxRate");
            Assert.Null(taxRate.Scope);
            Assert.Contains("0.21", taxRate.RefersTo);

            var range1 = names.Single(n => n.Name == "Range1");
            Assert.Equal("Data", range1.Scope);
            Assert.Contains("Data!$A$1:$B$1", range1.RefersTo);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ListDefinedNames_returns_empty_when_workbook_has_none()
    {
        var path = TestExcelWorkbooks.Create(_ => { });

        try
        {
            var names = new ExcelWorkbookService().ListDefinedNames(path);

            Assert.Empty(names);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ListDefinedNames_throws_file_not_found_for_missing_workbook()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"mcpoffice-missing-{Guid.NewGuid():N}.xlsx");

        var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().ListDefinedNames(missing));

        Assert.Contains("file_not_found", ex.Message);
    }
}
