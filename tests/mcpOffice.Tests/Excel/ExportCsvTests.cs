using DevExpress.Spreadsheet;
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel;

public class ExportCsvTests
{
    [Fact]
    public void Writes_used_range_with_mixed_types()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["C1"].Value = "Active";
            sheet.Cells["D1"].Value = "Joined";
            sheet.Cells["A2"].Value = "Ada";
            sheet.Cells["B2"].Value = 0.21;
            sheet.Cells["C2"].Value = true;
            sheet.Cells["D2"].Value = new DateTime(2026, 5, 7, 14, 30, 0);
            sheet.Cells["D2"].NumberFormat = "yyyy-mm-dd hh:mm:ss";
            sheet.Cells["A3"].Value = "Bob, Jr.";   // forces quoting
            sheet.Cells["B3"].Value = 1234567.89;
            sheet.Cells["C3"].Value = false;
            // D3 left blank
        });
        var output = TempPath(".csv");

        try
        {
            var result = new ExcelWorkbookService().ExportCsv(
                input,
                output,
                sheetName: "Data",
                sheetIndex: null,
                range: null,
                overwrite: false,
                maxRows: 1_048_576);

            Assert.Equal(output, result.OutputPath);
            Assert.Equal(3, result.RowCount);
            Assert.Equal(4, result.ColumnCount);
            Assert.True(result.BytesWritten > 0);

            var text = File.ReadAllText(output);
            Assert.Equal(
                "Name,Amount,Active,Joined\r\n" +
                "Ada,0.21,true,2026-05-07T14:30:00\r\n" +
                "\"Bob, Jr.\",1234567.89,false,",
                text);
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Range_parameter_slices_to_subset()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            for (var r = 1; r <= 5; r++)
            {
                for (var c = 0; c < 4; c++)
                {
                    sheet.Cells[r - 1, c].Value = $"r{r}c{c + 1}";
                }
            }
        });
        var output = TempPath(".csv");

        try
        {
            var result = new ExcelWorkbookService().ExportCsv(
                input, output, null, null, range: "A1:B3", overwrite: false, maxRows: 1_048_576);

            Assert.Equal(3, result.RowCount);
            Assert.Equal(2, result.ColumnCount);
            Assert.Equal(
                "r1c1,r1c2\r\nr2c1,r2c2\r\nr3c1,r3c2",
                File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Resolves_by_sheet_index_when_name_is_omitted()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Name = "First";
            workbook.Worksheets[0].Cells["A1"].Value = "one";
            workbook.Worksheets.Add("Second");
            workbook.Worksheets[1].Cells["A1"].Value = "two";
        });
        var output = TempPath(".csv");

        try
        {
            var result = new ExcelWorkbookService().ExportCsv(
                input, output, null, sheetIndex: 1, range: null, overwrite: false, maxRows: 1_048_576);

            Assert.Equal("two", File.ReadAllText(output));
            Assert.Equal(1, result.RowCount);
            Assert.Equal(1, result.ColumnCount);
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Throws_sheet_not_found_for_unknown_name()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Cells["A1"].Value = "x";
        });
        var output = TempPath(".csv");

        try
        {
            var ex = Assert.Throws<ModelContextProtocol.McpException>(() =>
                new ExcelWorkbookService().ExportCsv(
                    input, output, sheetName: "Missing", sheetIndex: null, range: null,
                    overwrite: false, maxRows: 1_048_576));

            Assert.Contains("[sheet_not_found]", ex.Message);
            Assert.Contains("Missing", ex.Message);
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Overwrite_false_throws_file_exists_when_output_exists()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Cells["A1"].Value = "x";
        });
        var output = TempPath(".csv");
        File.WriteAllText(output, "preexisting");

        try
        {
            var ex = Assert.Throws<ModelContextProtocol.McpException>(() =>
                new ExcelWorkbookService().ExportCsv(
                    input, output, null, null, null, overwrite: false, maxRows: 1_048_576));

            Assert.Contains("[file_exists]", ex.Message);
            Assert.Equal("preexisting", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Overwrite_true_replaces_existing_file()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Cells["A1"].Value = "fresh";
        });
        var output = TempPath(".csv");
        File.WriteAllText(output, "stale");

        try
        {
            new ExcelWorkbookService().ExportCsv(
                input, output, null, null, null, overwrite: true, maxRows: 1_048_576);

            Assert.Equal("fresh", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Creates_missing_output_directory()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Cells["A1"].Value = "x";
        });
        var subdir = Path.Combine(Path.GetTempPath(), $"mcpoffice-csv-{Guid.NewGuid():N}", "nested");
        var output = Path.Combine(subdir, "out.csv");

        try
        {
            new ExcelWorkbookService().ExportCsv(
                input, output, null, null, null, overwrite: false, maxRows: 1_048_576);

            Assert.True(File.Exists(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (Directory.Exists(subdir)) Directory.Delete(subdir, recursive: true);
        }
    }

    [Fact]
    public void Throws_range_too_large_when_row_count_exceeds_max_rows()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            for (var r = 1; r <= 10; r++)
            {
                sheet.Cells[r - 1, 0].Value = $"row{r}";
            }
        });
        var output = TempPath(".csv");

        try
        {
            var ex = Assert.Throws<ModelContextProtocol.McpException>(() =>
                new ExcelWorkbookService().ExportCsv(
                    input, output, null, null, null, overwrite: false, maxRows: 5));

            Assert.Contains("[range_too_large]", ex.Message);
            Assert.False(File.Exists(output), "no output file should be created when maxRows is exceeded");
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Formula_cell_emits_cached_value_not_formula_text()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = 40;
            sheet.Cells["B1"].Formula = "=A1+2";
            workbook.Calculate();
        });
        var output = TempPath(".csv");

        try
        {
            new ExcelWorkbookService().ExportCsv(
                input, output, null, null, null, overwrite: false, maxRows: 1_048_576);

            Assert.Equal("40,42", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Empty_cells_emit_empty_unquoted_fields()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = "x";
            // B1, C1 left blank
            sheet.Cells["D1"].Value = "y";
        });
        var output = TempPath(".csv");

        try
        {
            new ExcelWorkbookService().ExportCsv(
                input, output, null, null, range: "A1:D1", overwrite: false, maxRows: 1_048_576);

            Assert.Equal("x,,,y", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Output_is_locale_neutral_when_host_culture_is_nlNL()
    {
        var original = System.Globalization.CultureInfo.CurrentCulture;
        try
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("nl-NL");

            var input = TestExcelWorkbooks.Create(workbook =>
            {
                workbook.Worksheets[0].Cells["A1"].Value = 0.21;
                workbook.Worksheets[0].Cells["B1"].Value = 1234567.89;
            });
            var output = TempPath(".csv");

            try
            {
                new ExcelWorkbookService().ExportCsv(
                    input, output, null, null, null, overwrite: false, maxRows: 1_048_576);

                Assert.Equal("0.21,1234567.89", File.ReadAllText(output));
            }
            finally
            {
                if (File.Exists(input)) File.Delete(input);
                if (File.Exists(output)) File.Delete(output);
            }
        }
        finally
        {
            System.Globalization.CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void TrimTrailingEmptyRows_truncates_to_last_non_empty_row()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = "h1";
            sheet.Cells["B1"].Value = "h2";
            sheet.Cells["A2"].Value = "data";
            sheet.Cells["B2"].Value = 42;
            // C3..C10 left empty; an anchor cell extends the workbook used range
            sheet.Cells["Z99"].Value = "anchor";
        });
        var output = TempPath(".csv");

        try
        {
            var result = new ExcelWorkbookService().ExportCsv(
                input, output, null, null, range: "A1:B10",
                overwrite: false, maxRows: 1_048_576, trimTrailingEmptyRows: true);

            Assert.Equal(2, result.RowCount);
            Assert.Equal(2, result.ColumnCount);
            Assert.Equal("h1,h2\r\ndata,42", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void TrimTrailingEmptyRows_default_false_preserves_empty_rows()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = "h1";
            sheet.Cells["A2"].Value = "data";
            sheet.Cells["Z99"].Value = "anchor";
        });
        var output = TempPath(".csv");

        try
        {
            // No trimTrailingEmptyRows arg — default is false.
            var result = new ExcelWorkbookService().ExportCsv(
                input, output, null, null, range: "A1:A5",
                overwrite: false, maxRows: 1_048_576);

            Assert.Equal(5, result.RowCount);
            Assert.Equal("h1\r\ndata\r\n\r\n\r\n", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void TrimTrailingEmptyRows_treats_error_cells_as_empty()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = "h1";
            sheet.Cells["A2"].Value = "data";
            sheet.Cells["A3"].Formula = "=1/0";   // produces #DIV/0! after Calculate
            workbook.Calculate();
        });
        var output = TempPath(".csv");

        try
        {
            var result = new ExcelWorkbookService().ExportCsv(
                input, output, null, null, range: "A1:A3",
                overwrite: false, maxRows: 1_048_576, trimTrailingEmptyRows: true);

            Assert.Equal(2, result.RowCount);
            Assert.Equal("h1\r\ndata", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void TrimTrailingEmptyRows_keeps_data_when_last_row_has_data()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = "h1";
            sheet.Cells["A2"].Value = "data1";
            sheet.Cells["A3"].Value = "data2";
        });
        var output = TempPath(".csv");

        try
        {
            var result = new ExcelWorkbookService().ExportCsv(
                input, output, null, null, range: "A1:A3",
                overwrite: false, maxRows: 1_048_576, trimTrailingEmptyRows: true);

            Assert.Equal(3, result.RowCount);
            Assert.Equal("h1\r\ndata1\r\ndata2", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void TrimTrailingEmptyRows_all_empty_returns_zero_rows_zero_bytes()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Cells["Z99"].Value = "anchor"; // anchor outside the export range
        });
        var output = TempPath(".csv");

        try
        {
            var result = new ExcelWorkbookService().ExportCsv(
                input, output, null, null, range: "A1:B3",
                overwrite: false, maxRows: 1_048_576, trimTrailingEmptyRows: true);

            Assert.Equal(0, result.RowCount);
            Assert.Equal(2, result.ColumnCount);
            Assert.Equal(0, result.BytesWritten);
            Assert.Equal("", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    private static string TempPath(string ext) =>
        Path.Combine(Path.GetTempPath(), $"mcpoffice-csv-{Guid.NewGuid():N}{ext}");
}
