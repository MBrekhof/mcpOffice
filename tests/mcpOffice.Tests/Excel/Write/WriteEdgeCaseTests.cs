using System.Security.Cryptography;
using System.Text.Json;
using DevExpress.Spreadsheet;
using McpOffice.Models;
using McpOffice.Services.Excel;
using ModelContextProtocol;
using static McpOffice.Tests.Excel.Write.WriteTestSupport;

namespace McpOffice.Tests.Excel.Write;

/// <summary>Cases from the 2026-11-04 review of #2853, each reproduced before its fix.</summary>
public class WriteEdgeCaseTests
{
    private static string Input() => TestExcelWorkbooks.Create(wb =>
    {
        var ws = wb.Worksheets[0];
        ws.Name = "Data";
        ws.Cells["A1"].Value = 1;
        wb.Worksheets.Add("Other").Cells["A1"].Value = "other";
    });

    private static McpException Fails(string input, params ExcelOperation[] ops)
    {
        var output = TempPath(".xlsx");
        var ex = Assert.Throws<McpException>(() => Update(input, output, ops));
        Assert.False(File.Exists(output));
        return ex;
    }

    [Fact]
    public void Forward_references_inside_one_writeRange_get_their_values()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            Update(input, output, new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A1",
                Rows: Rows(["=B1*2", 10], ["=SUM(B1:B2)", 20], ["=A1+A2", "=A3*2"])));
            var data = Read(output, "Data");
            Assert.Equal(20d, data.Cell("A1").Value);
            Assert.Equal(30d, data.Cell("A2").Value);
            Assert.Equal(50d, data.Cell("A3").Value);     // depends on two written formulas
            Assert.Equal(100d, data.Cell("B3").Value);    // depends on a written formula that itself depends on written formulas
        }
        finally { Delete(input, output); }
    }

    [Theory]
    [InlineData("XFD1048576", 1, 2)]
    [InlineData("A1048576", 2, 1)]
    public void WriteRange_running_past_the_sheet_edge_is_invalid_operation(string startCell, int rows, int cols)
    {
        var input = Input();
        try
        {
            var block = Enumerable.Range(0, rows).Select(_ => Enumerable.Repeat<object?>(1, cols).ToArray()).ToArray();
            var ex = Fails(input, new ExcelOperation("writeRange", Sheet: "Data", StartCell: startCell, Rows: Rows(block)));
            Assert.StartsWith("[invalid_operation] operations[0] (writeRange)", ex.Message);
            Assert.Contains("past the last", ex.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void Deleting_the_only_visible_sheet_is_invalid_operation()
    {
        var input = TestExcelWorkbooks.Create(wb =>
        {
            wb.Worksheets[0].Name = "Shown";
            wb.Worksheets.Add("Hidden").Visible = false;
        });
        try
        {
            var ex = Fails(input, new ExcelOperation("deleteSheet", Sheet: "Shown"));
            Assert.StartsWith("[invalid_operation] operations[0] (deleteSheet): cannot delete 'Shown': it is the only visible sheet", ex.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void Insert_and_delete_beyond_the_sheet_are_invalid_operation()
    {
        var input = Input();
        try
        {
            var ex = Fails(input, new ExcelOperation("insertRows", Sheet: "Data", Row: 1048575, Count: 5));
            Assert.StartsWith("[invalid_operation] operations[0] (insertRows): row 1048575 + count 5 runs past", ex.Message);
            ex = Fails(input, new ExcelOperation("deleteRows", Sheet: "Data", Row: 1048575, Count: 5));
            Assert.StartsWith("[invalid_operation] operations[0] (deleteRows)", ex.Message);
            ex = Fails(input, new ExcelOperation("insertColumns", Sheet: "Data", Column: "A", Count: 16384));
            Assert.StartsWith("[invalid_operation] operations[0] (insertColumns): column A + count 16384 runs past", ex.Message);
            ex = Fails(input, new ExcelOperation("deleteColumns", Sheet: "Data", Column: "XFD", Count: 2));
            Assert.StartsWith("[invalid_operation] operations[0] (deleteColumns)", ex.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void Insert_that_would_push_data_off_the_sheet_is_invalid_operation()
    {
        var input = TestExcelWorkbooks.Create(wb =>
        {
            wb.Worksheets[0].Name = "Data";
            wb.Worksheets[0].Cells["A1048576"].Value = "last";
        });
        try
        {
            var ex = Fails(input, new ExcelOperation("insertRows", Sheet: "Data", Row: 1));
            Assert.StartsWith("[invalid_operation] operations[0] (insertRows)", ex.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void Union_and_cross_sheet_ranges_are_invalid_operation()
    {
        var input = Input();
        try
        {
            var ex = Fails(input, new ExcelOperation("autoFilter", Sheet: "Data", Range: "A1:B2,D1:E2"));
            Assert.StartsWith("[invalid_operation] operations[0] (autoFilter)", ex.Message);

            ex = Fails(input, new ExcelOperation("formatRange", Sheet: "Data", Range: "Other!A1", Bold: true));
            Assert.StartsWith("[invalid_operation] operations[0] (formatRange): range 'Other!A1' refers to another sheet", ex.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void Null_entries_are_invalid_operation_not_a_crash()
    {
        var input = Input();
        var path = TempPath(".xlsx");
        try
        {
            var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().CreateWorkbook(path, [null!], null, false));
            Assert.StartsWith("[invalid_operation] sheets[0]: sheet is null", ex.Message);

            ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().CreateWorkbook(path, [new ExcelSheetSpec("Data")], [null!], false));
            Assert.StartsWith("[invalid_operation] operations[0]: operation is null", ex.Message);

            ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().CreateWorkbook(path, [new ExcelSheetSpec("Data", Columns: [null!])], null, false));
            Assert.StartsWith("[invalid_operation] sheets[0]: columns[0] is null", ex.Message);

            ex = Fails(input, new ExcelOperation("addSheet", Name: "X"), null!);
            Assert.StartsWith("[invalid_operation] operations[1]: operation is null", ex.Message);
            Assert.False(File.Exists(path));
        }
        finally { Delete(input, path); }
    }

    [Fact]
    public void Device_prefixed_spelling_of_the_input_is_refused_and_input_untouched()
    {
        var input = Input();
        try
        {
            using var sha = SHA256.Create();
            var before = Convert.ToHexString(sha.ComputeHash(File.ReadAllBytes(input)));
            var spellings = new[] { @"\\?\" + input, @"\\.\" + input };
            foreach (var spelling in spellings)
            {
                var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().UpdateWorkbook(input, spelling, [new ExcelOperation("addSheet", Name: "X")], overwrite: true));
                Assert.StartsWith("[invalid_path]", ex.Message);
            }
            Assert.Equal(before, Convert.ToHexString(sha.ComputeHash(File.ReadAllBytes(input))));
        }
        finally { Delete(input); }
    }

    [Fact]
    public void Input_held_open_blocks_an_overwrite_through_an_unrecognised_alias()
    {
        // The string check cannot know every alias (junctions, 8.3 names). The OS can: the input is
        // held open without write sharing, so moving the new file over it fails instead of clobbering.
        var input = Input();
        try
        {
            var before = File.ReadAllBytes(input);
            using (new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var temp = TempPath(".xlsx");
                File.WriteAllBytes(temp, [1, 2, 3]);
                var ex = Assert.ThrowsAny<Exception>(() => File.Move(temp, input, overwrite: true));
                Assert.True(ex is IOException or UnauthorizedAccessException, ex.GetType().Name);
                Delete(temp);
            }
            Assert.Equal(before, File.ReadAllBytes(input));
        }
        finally { Delete(input); }
    }

    [Fact]
    public void Column_number_format_is_applied_before_autofit()
    {
        var path = TempPath(".xlsx");
        try
        {
            Create(path, new ExcelSheetSpec("Data", Rows(["Amount"], [1234567.89]),
                Columns: [new ExcelColumnSpec("A", NumberFormat: "#,##0.000000 \"euro\"")]));
            using var wb = Load(path);
            // "1,234,567.890000 euro" is 21 characters; the raw number auto-fits to about 10.
            Assert.True(wb.Worksheets[0].Columns["A"].WidthInCharacters >= 18, $"A is {wb.Worksheets[0].Columns["A"].WidthInCharacters}");
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Column_specs_apply_and_validate_without_data_rows()
    {
        var path = TempPath(".xlsx");
        try
        {
            Create(path, new ExcelSheetSpec("Empty", Columns: [new ExcelColumnSpec("B", Width: 33)]));
            using (var wb = Load(path)) Assert.Equal(33d, wb.Worksheets[0].Columns["B"].WidthInCharacters, 0.5);

            var ex = Assert.Throws<McpException>(() => Create(TempPath(".xlsx"),
                new ExcelSheetSpec("HeaderOnly", Rows(["a"]), Columns: [new ExcelColumnSpec("A", HorizontalAlign: "middle")])));
            Assert.StartsWith("[invalid_operation] sheets[0]: horizontalAlign 'middle'", ex.Message);
            ex = Assert.Throws<McpException>(() => Create(TempPath(".xlsx"),
                new ExcelSheetSpec("Bad", Columns: [new ExcelColumnSpec("A", Width: 300)])));
            Assert.StartsWith("[invalid_operation] sheets[0]: columns[0]: width 300", ex.Message);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void Localized_function_name_is_an_error_on_create_and_a_warning_on_update()
    {
        var path = TempPath(".xlsx");
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var ex = Assert.Throws<McpException>(() => Create(path, new ExcelSheetSpec("Data", Rows(["x", "y"], [1, "=SOM(A1,A2)"]))));
            Assert.StartsWith("[invalid_operation] sheets[0]: cell B2: formula '=SOM(A1,A2)' evaluates to #NAME?", ex.Message);
            Assert.False(File.Exists(path));

            var result = Update(input, output, new ExcelOperation("writeRange", Sheet: "Data", StartCell: "B1", Rows: Rows(["=GEMIDDELDE(A1:A3)"])));
            var warning = Assert.Single(result.Warnings);
            Assert.StartsWith("cell B1: formula '=GEMIDDELDE(A1:A3)' evaluates to #NAME?", warning);
            Assert.Contains("VBA", warning);
            Assert.True(File.Exists(output));
        }
        finally { Delete(path, input, output); }
    }

    [Fact]
    public void Formula_over_8192_chars_is_rejected()
    {
        var ex = Assert.Throws<McpException>(() => Create(TempPath(".xlsx"), new ExcelSheetSpec("Data", Rows(["=" + new string('1', 8192)]))));
        Assert.StartsWith("[invalid_operation] sheets[0]: cell A1: formula is 8193 characters", ex.Message);
    }

    [Fact]
    public void Failed_create_leaves_no_new_directory_behind()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"mcpoffice-write-nodir-{Guid.NewGuid():N}");
        try
        {
            Assert.Throws<McpException>(() => Create(Path.Combine(dir, "x.xlsx"), new ExcelSheetSpec("a/b")));
            Assert.False(Directory.Exists(dir));
        }
        finally { if (Directory.Exists(dir)) Directory.Delete(dir, true); }
    }

    // --- gaps the review listed ---

    [Fact]
    public void Update_overwrite_flag()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            Update(input, output, new ExcelOperation("addSheet", Name: "First"));
            var ex = Assert.Throws<McpException>(() => Update(input, output, new ExcelOperation("addSheet", Name: "Second")));
            Assert.StartsWith("[file_exists]", ex.Message);

            new ExcelWorkbookService().UpdateWorkbook(input, output, [new ExcelOperation("addSheet", Name: "Second")], overwrite: true);
            Assert.Equal(["Data", "Other", "Second"], new ExcelWorkbookService().ListSheets(output).Select(s => s.Name));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Update_cell_limit_counts_every_writeRange()
    {
        var input = Input();
        try
        {
            var block = Rows(Enumerable.Repeat(Enumerable.Repeat<object?>(1, 100).ToArray(), 501).ToArray());   // 50,100 cells
            var ex = Fails(input,
                new ExcelOperation("writeRange", Sheet: "Data", StartCell: "A1", Rows: block),
                new ExcelOperation("writeRange", Sheet: "Other", StartCell: "A1", Rows: block));
            Assert.StartsWith("[range_too_large]", ex.Message);
            Assert.Contains("100200", ex.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void FreezePanes_columns_only()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            Update(input, output, new ExcelOperation("freezePanes", Sheet: "Data", Cell: "C1"));
            var xml = SheetXml(output);
            Assert.Contains("xSplit=\"2\"", xml);
            Assert.DoesNotContain("ySplit", xml);
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void FormatRange_fontSize_and_borderScope_validation()
    {
        var input = Input();
        try
        {
            var ex = Fails(input, new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", FontSize: 0.5));
            Assert.StartsWith("[invalid_operation] operations[0] (formatRange): fontSize 0.5", ex.Message);
            ex = Fails(input, new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", FontSize: 500));
            Assert.StartsWith("[invalid_operation] operations[0] (formatRange): fontSize 500", ex.Message);
            ex = Fails(input, new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", BorderScope: "outline"));
            Assert.StartsWith("[invalid_operation] operations[0] (formatRange): borderScope needs border", ex.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void RenameSheet_rules()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var ex = Fails(input, new ExcelOperation("renameSheet", Sheet: "Data", NewName: "other"));
            Assert.Contains("already exists", ex.Message);
            ex = Fails(input, new ExcelOperation("renameSheet", Sheet: "Data", NewName: "a*b"));
            Assert.StartsWith("[invalid_operation] operations[0] (renameSheet): sheet name 'a*b'", ex.Message);
            ex = Fails(input, new ExcelOperation("renameSheet", Sheet: "Data"));
            Assert.StartsWith("[invalid_operation] operations[0] (renameSheet): newName is required", ex.Message);

            Update(input, output, new ExcelOperation("renameSheet", Sheet: "Data", NewName: "DATA"));   // case change of itself is fine
            Assert.Equal(["DATA", "Other"], new ExcelWorkbookService().ListSheets(output).Select(s => s.Name));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void AddSheet_position_bounds()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var ex = Fails(input, new ExcelOperation("addSheet", Name: "X", Position: 3));
            Assert.StartsWith("[invalid_operation] operations[0] (addSheet): position 3 is outside 0..2", ex.Message);
            ex = Fails(input, new ExcelOperation("addSheet", Name: "X", Position: -1));
            Assert.StartsWith("[invalid_operation] operations[0] (addSheet): position -1", ex.Message);

            Update(input, output, new ExcelOperation("addSheet", Name: "Last", Position: 2));
            Assert.Equal(["Data", "Other", "Last"], new ExcelWorkbookService().ListSheets(output).Select(s => s.Name));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Ragged_and_null_rows_are_fine()
    {
        var path = TempPath(".xlsx");
        try
        {
            var rows = new JsonElement[][]
            {
                Rows(["a", "b", "c"])[0],
                Rows([1])[0],
                null!,
                [],
                Rows([null, null, 3, 4])[0],
            };
            var result = Create(path, new ExcelSheetSpec("Data", rows));
            Assert.Equal(8, result.CellsWritten);
            Assert.Equal("A1:D5", result.Sheets[0].UsedRange);
            var data = Read(path);
            Assert.Equal(1d, data.Cell("A2").Value);
            Assert.Equal(4d, data.Cell("D5").Value);
            using var wb = Load(path);
            Assert.Equal("A1:D5", wb.Worksheets[0].AutoFilter.Range.GetReferenceA1());
        }
        finally { Delete(path); }
    }

    [Theory]
    [InlineData(".csv")]
    [InlineData(".xls")]
    public void Update_rejects_bad_input_extension_and_corrupt_input(string extension)
    {
        var input = TempPath(extension);
        var corrupt = TempPath(".xlsx");
        try
        {
            File.WriteAllText(input, "not a workbook");
            var ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("addSheet", Name: "X")));
            Assert.StartsWith("[unsupported_format] inputPath must be .xlsx or .xlsm", ex.Message);

            File.WriteAllText(corrupt, "not a workbook");
            ex = Assert.Throws<McpException>(() => Update(corrupt, TempPath(".xlsx"), new ExcelOperation("addSheet", Name: "X")));
            Assert.StartsWith("[parse_error]", ex.Message);
        }
        finally { Delete(input, corrupt); }
    }
}
