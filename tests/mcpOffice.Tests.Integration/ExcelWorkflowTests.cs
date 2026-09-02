using DevExpress.Spreadsheet;
using ModelContextProtocol.Protocol;
using SpreadsheetFormat = DevExpress.Spreadsheet.DocumentFormat;

namespace McpOffice.Tests.Integration;

public class ExcelWorkflowTests
{
    [Fact]
    public async Task List_sheets_via_stdio()
    {
        var path = TempPath(".xlsx");
        try
        {
            using (var workbook = new Workbook())
            {
                workbook.Worksheets[0].Name = "Data";
                workbook.Worksheets[0].Cells["A1"].Value = "Name";
                workbook.Worksheets[0].Cells["B2"].Value = 42;
                workbook.Worksheets.Add("Second");
                workbook.SaveDocument(path, SpreadsheetFormat.Xlsx);
            }

            await using var harness = await ServerHarness.StartAsync();
            var result = await harness.Client.CallToolAsync(
                "excel_list_sheets",
                new Dictionary<string, object?> { ["path"] = path });
            var text = result.Content.OfType<TextContentBlock>().Single().Text;

            Assert.Contains("\"name\":\"Data\"", text);
            Assert.Contains("\"usedRange\":\"A1:B2\"", text);
            Assert.Contains("\"name\":\"Second\"", text);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task Read_sheet_via_stdio()
    {
        var path = TempPath(".xlsx");
        try
        {
            using (var workbook = new Workbook())
            {
                var sheet = workbook.Worksheets[0];
                sheet.Name = "Data";
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "Ada";
                sheet.Cells["B2"].Value = 40;
                sheet.Cells["C2"].Formula = "=B2+2";
                workbook.Calculate();
                workbook.SaveDocument(path, SpreadsheetFormat.Xlsx);
            }

            await using var harness = await ServerHarness.StartAsync();
            var result = await harness.Client.CallToolAsync(
                "excel_read_sheet",
                new Dictionary<string, object?>
                {
                    ["path"] = path,
                    ["sheetName"] = "Data",
                    ["range"] = "A1:C2",
                    ["includeFormulas"] = true,
                    ["includeFormats"] = false,
                    ["maxCells"] = 10
                });
            var text = result.Content.OfType<TextContentBlock>().Single().Text;

            Assert.Contains("\"sheet\":\"Data\"", text);
            Assert.Contains("\"range\":\"A1:C2\"", text);
            Assert.Contains("\"Ada\"", text);
            Assert.Contains("\"value\":42", text);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task Extract_vba_via_stdio_returns_modules()
    {
        var fixture = ResolveFixturePath("sample-with-macros.xlsm");
        if (!File.Exists(fixture))
        {
            // Hand-authored fixture not yet on disk; covered by VbaProjectReaderTests
            // (synthetic) until it lands. See plan Task 11.
            return;
        }

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_extract_vba",
            new Dictionary<string, object?> { ["path"] = fixture });

        var text = result.Content.OfType<TextContentBlock>().Single().Text;

        Assert.Contains("\"hasVbaProject\":true", text);
        Assert.Contains("\"name\":\"Module1\"", text);
        Assert.Contains("\"kind\":\"standardModule\"", text);
        Assert.Contains("Sub Hello", text);
    }

    [Fact]
    public async Task Extract_vba_via_stdio_returns_empty_for_xlsx_without_macros()
    {
        var path = TempPath(".xlsx");
        try
        {
            using (var workbook = new Workbook())
            {
                workbook.Worksheets[0].Cells["A1"].Value = "x";
                workbook.SaveDocument(path, SpreadsheetFormat.Xlsx);
            }

            await using var harness = await ServerHarness.StartAsync();
            var result = await harness.Client.CallToolAsync(
                "excel_extract_vba",
                new Dictionary<string, object?> { ["path"] = path });

            var text = result.Content.OfType<TextContentBlock>().Single().Text;

            Assert.Contains("\"hasVbaProject\":false", text);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task List_vba_entry_points_via_stdio()
    {
        var fixture = ResolveFixturePath("synthetic-vba.xlsm");
        if (!File.Exists(fixture)) return;

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_list_vba_entry_points",
            new Dictionary<string, object?> { ["path"] = fixture });

        var text = result.Content.OfType<TextContentBlock>().Single().Text;

        Assert.Contains("\"hasVbaProject\":true", text);
        Assert.Contains("\"kind\":\"eventHandler\"", text);
        Assert.Contains("ThisWorkbook.Workbook_Open", text);
        Assert.Contains("\"unreachable\":[", text);
        Assert.Contains("Class1.Greet", text);   // nothing calls the class method: dead code
    }

    [Fact]
    public async Task Map_vba_sheet_access_via_stdio()
    {
        var fixture = ResolveFixturePath("synthetic-vba.xlsm");
        if (!File.Exists(fixture)) return;

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_map_vba_sheet_access",
            new Dictionary<string, object?> { ["path"] = fixture });

        var text = result.Content.OfType<TextContentBlock>().Single().Text;

        Assert.Contains("\"hasVbaProject\":true", text);
        Assert.Contains("\"sheetAccess\":[", text);
        Assert.Contains("\"sheets\":[", text);
    }

    [Fact]
    public async Task Compare_vba_corpus_via_stdio()
    {
        var a = ResolveFixturePath("synthetic-vba.xlsm");
        var b = ResolveFixturePath("sample-with-macros.xlsm");
        if (!File.Exists(a) || !File.Exists(b)) return;

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_compare_vba_corpus",
            new Dictionary<string, object?> { ["paths"] = new[] { a, b } });

        var text = result.Content.OfType<TextContentBlock>().Single().Text;

        Assert.Contains("\"workbookCount\":2", text);
        Assert.Contains("\"sharedProcedures\":[", text);
    }

    [Fact]
    public async Task Analyze_vba_via_stdio_returns_summary()
    {
        var fixture = ResolveFixturePath("sample-with-macros.xlsm");
        if (!File.Exists(fixture)) return;  // synthetic fixture optional; same skip pattern as Extract_vba_via_stdio

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_analyze_vba",
            new Dictionary<string, object?>
            {
                ["path"] = fixture,
                ["includeProcedures"] = true,
                ["includeCallGraph"] = true,
                ["includeReferences"] = true
            });

        var text = result.Content.OfType<TextContentBlock>().Single().Text;
        Assert.Contains("\"hasVbaProject\":true", text);
        Assert.Contains("\"summary\":", text);
        Assert.Contains("\"modules\":", text);
    }

    [Fact]
    public async Task Render_vba_callgraph_via_stdio_returns_mermaid()
    {
        var fixture = ResolveFixturePath("sample-with-macros.xlsm");
        if (!File.Exists(fixture)) return;

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_render_vba_callgraph",
            new Dictionary<string, object?>
            {
                ["path"] = fixture,
                ["format"] = "mermaid",
                ["layout"] = "flat"
            });

        var text = result.Content.OfType<TextContentBlock>().Single().Text;
        Assert.NotEmpty(text);
        Assert.Contains("flowchart TD", text);
    }

    [Fact]
    public async Task Render_vba_callgraph_returns_empty_flowchart_for_xlsx_without_macros()
    {
        var path = TempPath(".xlsx");
        try
        {
            using (var workbook = new Workbook())
            {
                workbook.Worksheets[0].Cells["A1"].Value = "x";
                workbook.SaveDocument(path, SpreadsheetFormat.Xlsx);
            }

            await using var harness = await ServerHarness.StartAsync();
            var result = await harness.Client.CallToolAsync(
                "excel_render_vba_callgraph",
                new Dictionary<string, object?> { ["path"] = path });

            var text = result.Content.OfType<TextContentBlock>().Single().Text;
            Assert.Contains("flowchart TD", text);
            Assert.DoesNotContain("subgraph", text);
            Assert.DoesNotContain("-->", text);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Suggests_vba_conversion_via_stdio()
    {
        var fixture = ResolveFixturePath("synthetic-vba.xlsm");

        await using var harness = await ServerHarness.StartAsync();
        var response = await harness.Client.CallToolAsync(
            "excel_suggest_vba_conversion",
            new Dictionary<string, object?>
            {
                ["path"] = fixture,
                ["targetParadigm"] = "classLibrary"
            });

        var text = response.Content
            .OfType<TextContentBlock>()
            .Single().Text;

        // Loose JSON-shape assertions — the unit tests cover the matrix; this just
        // verifies the JSON-RPC layer doesn't drop fields.
        Assert.Contains("\"summary\"", text);
        Assert.Contains("\"procedureHints\"", text);
        Assert.Contains("\"moduleCoupling\"", text);
        Assert.Contains("\"couplingPairs\"", text);
        Assert.Contains("\"csharpSuggestion\"", text);
        Assert.Contains("\"targetParadigm\":\"classLibrary\"", text);
    }

    [Fact]
    public async Task Export_csv_via_stdio()
    {
        var input = TempPath(".xlsx");
        var output = TempPath(".csv");
        try
        {
            using (var workbook = new Workbook())
            {
                var sheet = workbook.Worksheets[0];
                sheet.Name = "Data";
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "Ada";
                sheet.Cells["B2"].Value = 0.21;
                workbook.SaveDocument(input, SpreadsheetFormat.Xlsx);
            }

            await using var harness = await ServerHarness.StartAsync();
            var result = await harness.Client.CallToolAsync(
                "excel_export_csv",
                new Dictionary<string, object?>
                {
                    ["path"] = input,
                    ["outputPath"] = output,
                    ["sheetName"] = "Data",
                });
            var text = result.Content.OfType<TextContentBlock>().Single().Text;

            Assert.Contains("\"rowCount\":2", text);
            Assert.Contains("\"columnCount\":2", text);
            Assert.True(File.Exists(output));
            Assert.Equal("Name,Amount\r\nAda,0.21", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input))  File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    private static string ResolveFixturePath(string name)
    {
        var asmDir = Path.GetDirectoryName(typeof(ExcelWorkflowTests).Assembly.Location)!;
        var dir = new DirectoryInfo(asmDir);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "mcpOffice.sln")))
            dir = dir.Parent;
        return Path.Combine(dir!.FullName, "tests", "fixtures", name);
    }

    private static string TempPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"mcpoffice-excel-integration-{Guid.NewGuid():N}{extension}");
}
