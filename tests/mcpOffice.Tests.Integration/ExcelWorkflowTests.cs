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

    private static string TempPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"mcpoffice-excel-integration-{Guid.NewGuid():N}{extension}");
}
