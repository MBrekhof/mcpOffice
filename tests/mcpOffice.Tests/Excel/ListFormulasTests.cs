using McpOffice.Services.Excel;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel;

public class ListFormulasTests
{
    [Fact]
    public void ListFormulas_returns_formula_cells_across_sheets()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            var data = workbook.Worksheets[0];
            data.Name = "Data";
            data.Cells["A1"].Value = 10;
            data.Cells["A2"].Value = 20;
            data.Cells["A3"].Formula = "=SUM(A1:A2)";

            var summary = workbook.Worksheets.Add("Summary");
            summary.Cells["B1"].Formula = "=Data!A3*2";
        });

        try
        {
            var formulas = new ExcelWorkbookService().ListFormulas(
                path,
                sheetName: null,
                includeValues: true,
                maxFormulas: 1000);

            Assert.Equal(2, formulas.Count);

            var sumFormula = formulas.Single(f => f.Sheet == "Data");
            Assert.Equal("A3", sumFormula.Address);
            Assert.Equal("=SUM(A1:A2)", sumFormula.Formula);
            Assert.Equal(30d, sumFormula.Value);
            Assert.Equal("number", sumFormula.ValueType);

            var summaryFormula = formulas.Single(f => f.Sheet == "Summary");
            Assert.Equal("B1", summaryFormula.Address);
            Assert.Contains("Data!A3", summaryFormula.Formula);
            Assert.Equal(60d, summaryFormula.Value);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ListFormulas_filters_by_sheet_name()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Name = "Data";
            workbook.Worksheets[0].Cells["A1"].Formula = "=1+1";
            var other = workbook.Worksheets.Add("Other");
            other.Cells["A1"].Formula = "=2+2";
        });

        try
        {
            var formulas = new ExcelWorkbookService().ListFormulas(
                path,
                sheetName: "Other",
                includeValues: false,
                maxFormulas: 1000);

            var only = Assert.Single(formulas);
            Assert.Equal("Other", only.Sheet);
            Assert.Null(only.Value);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ListFormulas_throws_range_too_large_when_count_exceeds_max()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            for (var i = 1; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].Formula = $"=ROW()+{i}";
            }
        });

        try
        {
            var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().ListFormulas(
                path,
                sheetName: null,
                includeValues: false,
                maxFormulas: 2));

            Assert.Contains("range_too_large", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ListFormulas_throws_sheet_not_found_for_unknown_name()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Cells["A1"].Formula = "=1+1";
        });

        try
        {
            var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().ListFormulas(
                path,
                sheetName: "Missing",
                includeValues: false,
                maxFormulas: 1000));

            Assert.Contains("sheet_not_found", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
