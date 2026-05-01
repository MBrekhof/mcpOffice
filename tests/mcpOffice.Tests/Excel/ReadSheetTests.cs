using McpOffice.Services.Excel;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel;

public class ReadSheetTests
{
    [Fact]
    public void ReadSheet_returns_rows_cells_values_and_formulas()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Ada";
            sheet.Cells["B2"].Value = 40;
            sheet.Cells["C2"].Formula = "=B2+2";
            workbook.Calculate();
        });

        try
        {
            var result = new ExcelWorkbookService().ReadSheet(
                path,
                sheetName: "Data",
                sheetIndex: null,
                range: "A1:C2",
                includeFormulas: true,
                includeFormats: true,
                maxCells: 10);

            Assert.Equal("Data", result.Sheet);
            Assert.Equal("A1:C2", result.Range);
            Assert.False(result.Truncated);
            Assert.Equal(2, result.Rows.Count);
            Assert.Equal("Name", result.Rows[0][0]);
            Assert.Equal("Ada", result.Rows[1][0]);
            Assert.Equal(40d, result.Rows[1][1]);

            var formulaCell = result.Cells.Single(c => c.Address == "C2");
            Assert.Equal("=B2+2", formulaCell.Formula);
            Assert.Equal(42d, formulaCell.Value);
            Assert.Equal("number", formulaCell.ValueType);
            Assert.NotNull(formulaCell.NumberFormat);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ReadSheet_uses_sheet_index_when_name_is_omitted()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Name = "First";
            workbook.Worksheets[0].Cells["A1"].Value = "one";
            var second = workbook.Worksheets.Add("Second");
            second.Cells["A1"].Value = "two";
        });

        try
        {
            var result = new ExcelWorkbookService().ReadSheet(
                path,
                sheetName: null,
                sheetIndex: 1,
                range: null,
                includeFormulas: true,
                includeFormats: false,
                maxCells: 10);

            Assert.Equal("Second", result.Sheet);
            Assert.Equal("two", result.Rows[0][0]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ReadSheet_throws_range_too_large_when_cell_count_exceeds_max()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Cells["A1"].Value = "a";
            workbook.Worksheets[0].Cells["B2"].Value = "b";
        });

        try
        {
            var ex = Assert.Throws<McpException>(() =>
                new ExcelWorkbookService().ReadSheet(
                    path,
                    sheetName: null,
                    sheetIndex: null,
                    range: "A1:B2",
                    includeFormulas: true,
                    includeFormats: false,
                    maxCells: 3));

            Assert.Contains("range_too_large", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ReadSheet_throws_sheet_not_found_for_unknown_name()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Name = "Data";
        });

        try
        {
            var ex = Assert.Throws<McpException>(() =>
                new ExcelWorkbookService().ReadSheet(
                    path,
                    sheetName: "Missing",
                    sheetIndex: null,
                    range: null,
                    includeFormulas: true,
                    includeFormats: false,
                    maxCells: 10));

            Assert.Contains("sheet_not_found", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
