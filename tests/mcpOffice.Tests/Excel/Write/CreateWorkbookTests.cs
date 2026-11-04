using System.Globalization;
using System.Text.Json;
using DevExpress.Spreadsheet;
using McpOffice.Models;
using McpOffice.Services.Excel;
using ModelContextProtocol;
using static McpOffice.Tests.Excel.Write.WriteTestSupport;

namespace McpOffice.Tests.Excel.Write;

public class CreateWorkbookTests
{
    private static ExcelSheetSpec Typed(string name = "Data", bool header = true, ExcelColumnSpec[]? columns = null) =>
        new(name, Rows(["Name", "Amount", "Date", "Total"], ["Ada", 40, "2026-11-04", "=B2*2"]), header, columns);

    [Fact]
    public void Create_writes_typed_values_readable_back()
    {
        var path = TempPath(".xlsx");
        try
        {
            var result = Create(path, Typed());

            Assert.Equal(path, result.OutputPath);
            Assert.Equal([new ExcelWrittenSheet("Data", "A1:D2")], result.Sheets);
            Assert.Equal(8, result.CellsWritten);
            Assert.Equal(0, result.OperationsApplied);
            Assert.Empty(result.Warnings);

            var data = Read(path);
            Assert.Equal("text", data.Cell("A2").ValueType);
            Assert.Equal("number", data.Cell("B2").ValueType);
            Assert.Equal(40d, data.Cell("B2").Value);
            Assert.Equal("datetime", data.Cell("C2").ValueType);
            Assert.Equal(new DateTime(2026, 11, 4), data.Cell("C2").Value);
            Assert.Equal("yyyy-mm-dd", data.Cell("C2").NumberFormat);
            Assert.Equal("number", data.Cell("D2").ValueType);
            Assert.Equal(80d, data.Cell("D2").Value);          // create recalculates
            Assert.Equal("=B2*2", data.Cell("D2").Formula);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Create_applies_header_defaults()
    {
        var path = TempPath(".xlsx");
        try
        {
            Create(path, Typed());

            using var wb = Load(path);
            var ws = wb.Worksheets[0];
            Assert.True(ws.Cells["A1"].Font.Bold);
            Assert.True(ws.Cells["D1"].Font.Bold);
            Assert.False(ws.Cells["A2"].Font.Bold);
            Assert.Equal(0xD9D9D9, ws.Cells["A1"].Fill.BackgroundColor.ToArgb() & 0xFFFFFF);
            Assert.Equal(BorderLineStyle.Thin, ws.Cells["A1"].Borders.BottomBorder.LineStyle);
            Assert.True(ws.AutoFilter.Enabled);
            Assert.Equal("A1:D2", ws.AutoFilter.Range.GetReferenceA1());

            var xml = SheetXml(path);
            Assert.Contains("state=\"frozen\"", xml);
            Assert.Contains("ySplit=\"1\"", xml);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Create_without_header_has_no_bold_freeze_or_filter()
    {
        var path = TempPath(".xlsx");
        try
        {
            Create(path, Typed(header: false));

            using var wb = Load(path);
            var ws = wb.Worksheets[0];
            Assert.False(ws.Cells["A1"].Font.Bold);
            Assert.False(ws.AutoFilter.Enabled);
            Assert.DoesNotContain("state=\"frozen\"", SheetXml(path));
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Create_autofit_caps_at_60_chars()
    {
        var path = TempPath(".xlsx");
        try
        {
            Create(path, new ExcelSheetSpec("Data", Rows(["Short", "Long"], ["ab", new string('x', 200)])));

            using var wb = Load(path);
            var ws = wb.Worksheets[0];
            Assert.True(ws.Columns["B"].WidthInCharacters <= 60, $"B is {ws.Columns["B"].WidthInCharacters}");
            Assert.True(ws.Columns["B"].WidthInCharacters >= 50, $"B is {ws.Columns["B"].WidthInCharacters}");
            Assert.True(ws.Columns["A"].WidthInCharacters < 20, $"A is {ws.Columns["A"].WidthInCharacters}");
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Create_applies_column_specs_to_data_cells_not_header()
    {
        var path = TempPath(".xlsx");
        try
        {
            Create(path, Typed(columns: [new ExcelColumnSpec("B", NumberFormat: "0.00", Width: 12, HorizontalAlign: "right")]));

            using var wb = Load(path);
            var ws = wb.Worksheets[0];
            Assert.Equal("0.00", ws.Cells["B2"].NumberFormat);
            Assert.NotEqual("0.00", ws.Cells["B1"].NumberFormat);
            Assert.Equal(SpreadsheetHorizontalAlignment.Right, ws.Cells["B2"].Alignment.Horizontal);
            Assert.NotEqual(SpreadsheetHorizontalAlignment.Right, ws.Cells["B1"].Alignment.Horizontal);
            Assert.Equal(12d, ws.Columns["B"].WidthInCharacters, 0.5);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Create_multiple_sheets_in_order_and_empty_sheet()
    {
        var path = TempPath(".xlsx");
        try
        {
            var result = Create(path,
                new ExcelSheetSpec("First", Rows(["a"])),
                new ExcelSheetSpec("Empty"),
                new ExcelSheetSpec("Third", Rows(["c"]), Header: false));

            Assert.Equal(["First", "Empty", "Third"], result.Sheets.Select(s => s.Name));
            Assert.Equal(["First", "Empty", "Third"], new ExcelWorkbookService().ListSheets(path).Select(s => s.Name));
            Assert.Equal("c", Read(path, "Third").Cell("A1").Value);
        }
        finally { Delete(path); }
    }

    [Theory]
    [InlineData(".xlsm")]
    [InlineData(".csv")]
    public void Create_rejects_non_xlsx_path(string extension)
    {
        var path = TempPath(extension);
        var ex = Assert.Throws<McpException>(() => Create(path, Typed()));
        Assert.StartsWith("[unsupported_format]", ex.Message);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Create_existing_file_without_overwrite_is_file_exists()
    {
        var path = TempPath(".xlsx");
        try
        {
            Create(path, Typed());
            var ex = Assert.Throws<McpException>(() => Create(path, Typed()));
            Assert.StartsWith("[file_exists]", ex.Message);

            new ExcelWorkbookService().CreateWorkbook(path, [new ExcelSheetSpec("Again", Rows(["z"]))], null, overwrite: true);
            Assert.Equal(["Again"], new ExcelWorkbookService().ListSheets(path).Select(s => s.Name));
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Duplicate_sheet_name_case_insensitive_is_invalid_operation()
    {
        var path = TempPath(".xlsx");
        var ex = Assert.Throws<McpException>(() => Create(path, new ExcelSheetSpec("Data"), new ExcelSheetSpec("data")));
        Assert.StartsWith("[invalid_operation] sheets[1]", ex.Message);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Invalid_sheet_name_is_invalid_operation()
    {
        var path = TempPath(".xlsx");
        var ex = Assert.Throws<McpException>(() => Create(path, new ExcelSheetSpec("a/b")));
        Assert.StartsWith("[invalid_operation] sheets[0]", ex.Message);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Formula_with_semicolon_separator_is_invalid_operation()
    {
        var path = TempPath(".xlsx");
        var ex = Assert.Throws<McpException>(() =>
            Create(path, new ExcelSheetSpec("Data", Rows(["x", "y"], [1, "=SUM(A1;A2)"]))));
        Assert.StartsWith("[invalid_operation] sheets[0]: cell B2", ex.Message);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Object_cell_value_is_invalid_operation_naming_the_cell()
    {
        var path = TempPath(".xlsx");
        var rows = new[] { new[] { JsonSerializer.SerializeToElement("ok"), JsonSerializer.SerializeToElement(new { a = 1 }) } };
        var ex = Assert.Throws<McpException>(() => Create(path, new ExcelSheetSpec("Data", rows)));
        Assert.StartsWith("[invalid_operation] sheets[0]: cell B1", ex.Message);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Over_100000_cells_is_range_too_large_and_writes_nothing()
    {
        var path = TempPath(".xlsx");
        var row = Enumerable.Repeat<object?>(1, 101).ToArray();
        var rows = Enumerable.Repeat(row, 1000).ToArray();   // 101,000 cells
        var ex = Assert.Throws<McpException>(() => Create(path, new ExcelSheetSpec("Data", Rows(rows))));
        Assert.StartsWith("[range_too_large]", ex.Message);
        Assert.Contains("101000", ex.Message);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Create_is_culture_independent()
    {
        var path = TempPath(".xlsx");
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("nl-NL");
            Create(path, new ExcelSheetSpec("Data", Rows(["Amount", "Total"], [12.5, "=A2*1.5"])));
            var data = Read(path);
            Assert.Equal(12.5d, data.Cell("A2").Value);
            Assert.Equal(18.75d, data.Cell("B2").Value);
            Assert.Equal("=A2*1.5", data.Cell("B2").Formula);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
            Delete(path);
        }
    }

    [Fact]
    public void Create_applies_operations_after_sheets()
    {
        var path = TempPath(".xlsx");
        try
        {
            var result = new ExcelWorkbookService().CreateWorkbook(
                path,
                [new ExcelSheetSpec("Data", Rows(["a", "b"], [1, 2]))],
                [
                    new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A3", Rows: Rows(["Total", "=SUM(B2:B2)"])),
                    new ExcelOperation("formatRange", Sheet: "Data", Range: "A3:B3", Bold: true),
                ],
                overwrite: false);

            Assert.Equal(6, result.CellsWritten);
            Assert.Equal(2, result.OperationsApplied);
            Assert.Equal(2d, Read(path).Cell("B3").Value);
            using var wb = Load(path);
            Assert.True(wb.Worksheets[0].Cells["A3"].Font.Bold);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Empty_sheets_list_is_invalid_operation()
    {
        var ex = Assert.Throws<McpException>(() => Create(TempPath(".xlsx")));
        Assert.StartsWith("[invalid_operation] sheets", ex.Message);
    }
}
