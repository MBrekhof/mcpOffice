using System.Security.Cryptography;
using McpOffice.Models;
using McpOffice.Services.Excel;
using ModelContextProtocol;
using static McpOffice.Tests.Excel.Write.WriteTestSupport;

namespace McpOffice.Tests.Excel.Write;

public class UpdateWorkbookTests
{
    /// <summary>A1=40, B1 "=A1*2" with cached 80, sheet "Data"; second sheet "Other".</summary>
    private static string Input() => TestExcelWorkbooks.Create(wb =>
    {
        var ws = wb.Worksheets[0];
        ws.Name = "Data";
        ws.Cells["A1"].Value = 40;
        ws.Cells["B1"].Formula = "=A1*2";
        ws.Cells["A1"].Font.Bold = true;
        wb.Worksheets.Add("Other");
        wb.Calculate();
    });

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    [Fact]
    public void Update_writes_new_version_and_leaves_input_byte_identical()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var before = Sha256(input);
            var result = Update(input, output, new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A2", Rows: Rows(["new", 7])));

            Assert.Equal(before, Sha256(input));
            Assert.Equal(output, result.OutputPath);
            Assert.Equal(2, result.CellsWritten);
            Assert.Equal(1, result.OperationsApplied);
            Assert.Empty(result.Warnings);
            Assert.Equal("new", Read(output, "Data").Cell("A2").Value);
            Assert.Equal(7d, Read(output, "Data").Cell("B2").Value);
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Update_output_equal_to_input_is_invalid_path()
    {
        var input = Input();
        try
        {
            var op = new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A2", Rows: Rows([1]));
            foreach (var same in new[] { input, input.ToUpperInvariant(), input.Replace(@"\", @"/") })
            {
                var ex = Assert.Throws<McpException>(() => Update(input, same, op));
                Assert.StartsWith("[invalid_path]", ex.Message);
            }
        }
        finally { Delete(input); }
    }

    [Fact]
    public void Update_does_not_recalculate_untouched_formulas()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            Update(input, output, new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A1", Rows: Rows([50])));

            var data = Read(output, "Data");
            Assert.Equal(50d, data.Cell("A1").Value);
            Assert.Equal(80d, data.Cell("B1").Value);      // cached value untouched; Excel recalculates on open
            Assert.Equal("=A1*2", data.Cell("B1").Formula);
            Assert.Contains("fullCalcOnLoad=\"1\"", WorkbookXml(output));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Formula_written_by_update_is_evaluated_on_write_while_untouched_cached_values_stay()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            Update(input, output,
                new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A2", Rows: Rows([7, "=A2*2"])),
                new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A1", Rows: Rows([50])));

            var data = Read(output, "Data");
            Assert.Equal(14d, data.Cell("B2").Value);     // DevExpress evaluates a formula when it is assigned
            Assert.Equal(80d, data.Cell("B1").Value);     // but does not touch the cached value of an existing one
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Failing_operation_writes_nothing()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var ex = Assert.Throws<McpException>(() => Update(input, output,
                new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A2", Rows: Rows([1])),
                new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A3", Rows: Rows([2])),
                new ExcelOperation("formatRange", Sheet: "Missing", Range: "A1", Bold: true)));

            Assert.StartsWith("[sheet_not_found]", ex.Message);
            Assert.Contains("operations[2] (formatRange)", ex.Message);
            Assert.False(File.Exists(output));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Unknown_op_is_invalid_operation_listing_valid_ops()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var ex = Assert.Throws<McpException>(() => Update(input, output, new ExcelOperation("paintItRed", Sheet: "Data")));
            Assert.StartsWith("[invalid_operation] operations[0] (paintItRed): unknown op", ex.Message);
            Assert.Contains("writeRange", ex.Message);
            Assert.Contains("autoFilter", ex.Message);
            Assert.False(File.Exists(output));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Sheet_operations()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var result = Update(input, output,
                new ExcelOperation("addSheet", Name: "Appended"),
                new ExcelOperation("addSheet", Name: "First", Position: 0),
                new ExcelOperation("renameSheet", Sheet: "other", NewName: "Renamed"),
                new ExcelOperation("deleteSheet", Sheet: "Appended"));

            Assert.Equal(["First", "Data", "Renamed"], result.Sheets.Select(s => s.Name));
            Assert.Equal(["First", "Data", "Renamed"], new ExcelWorkbookService().ListSheets(output).Select(s => s.Name));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Deleting_the_last_sheet_and_duplicate_addSheet_are_invalid_operation()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var ex = Assert.Throws<McpException>(() => Update(input, output,
                new ExcelOperation("deleteSheet", Sheet: "Other"),
                new ExcelOperation("deleteSheet", Sheet: "Data")));
            Assert.StartsWith("[invalid_operation] operations[1] (deleteSheet)", ex.Message);

            ex = Assert.Throws<McpException>(() => Update(input, output, new ExcelOperation("addSheet", Name: "DATA")));
            Assert.StartsWith("[invalid_operation] operations[0] (addSheet)", ex.Message);
            Assert.Contains("already exists", ex.Message);

            ex = Assert.Throws<McpException>(() => Update(input, output, new ExcelOperation("addSheet", Name: "bad:name")));
            Assert.StartsWith("[invalid_operation] operations[0] (addSheet)", ex.Message);
            Assert.False(File.Exists(output));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void WriteRange_null_clears_cell_and_clearRange_modes()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            Update(input, output,
                new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A1", Rows: Rows([null, "kept"])),
                new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A3", Rows: Rows(["c", "f", "a"])),
                new ExcelOperation("formatRange", Sheet: "Data", Range: "A3:C3", Bold: true),
                new ExcelOperation("clearRange", Sheet: "Data", Range: "A3", What: "contents"),
                new ExcelOperation("clearRange", Sheet: "Data", Range: "B3", What: "formats"),
                new ExcelOperation("clearRange", Sheet: "Data", Range: "C3", What: "all"));

            var data = Read(output, "Data", "A1:C3");
            Assert.Equal("empty", data.Cell("A1").ValueType);
            Assert.Equal("kept", data.Cell("B1").Value);
            Assert.Equal("empty", data.Cell("A3").ValueType);
            Assert.Equal("f", data.Cell("B3").Value);
            Assert.Equal("empty", data.Cell("C3").ValueType);

            using var wb = Load(output);
            var ws = wb.Worksheets["Data"];
            Assert.True(ws.Cells["A3"].Font.Bold);     // contents cleared, format kept
            Assert.False(ws.Cells["B3"].Font.Bold);    // format cleared, value kept
            Assert.False(ws.Cells["C3"].Font.Bold);

            var ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"),
                new ExcelOperation("clearRange", Sheet: "Data", Range: "A1", What: "values")));
            Assert.StartsWith("[invalid_operation] operations[0] (clearRange): what 'values'", ex.Message);
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Update_creates_output_directory_and_missing_input_is_file_not_found()
    {
        var input = Input();
        var dir = Path.Combine(Path.GetTempPath(), $"mcpoffice-write-dir-{Guid.NewGuid():N}");
        var output = Path.Combine(dir, "sub", "v2.xlsx");
        try
        {
            Update(input, output, new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A2", Rows: Rows([1])));
            Assert.True(File.Exists(output));

            var missing = TempPath(".xlsx");
            var ex = Assert.Throws<McpException>(() => Update(missing, TempPath(".xlsx"), new ExcelOperation("addSheet", Name: "X")));
            Assert.StartsWith("[file_not_found]", ex.Message);
        }
        finally
        {
            Delete(input);
            if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Update_rejects_empty_operations_and_bad_extensions()
    {
        var input = Input();
        try
        {
            var ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx")));
            Assert.StartsWith("[invalid_operation] operations", ex.Message);

            ex = Assert.Throws<McpException>(() => Update(input, TempPath(".csv"), new ExcelOperation("addSheet", Name: "X")));
            Assert.StartsWith("[unsupported_format]", ex.Message);
        }
        finally { Delete(input); }
    }
}
