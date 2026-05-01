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
