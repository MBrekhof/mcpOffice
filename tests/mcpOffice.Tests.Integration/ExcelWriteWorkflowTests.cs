using System.Security.Cryptography;
using System.Text.Json;
using ModelContextProtocol.Protocol;

namespace McpOffice.Tests.Integration;

/// <summary>
/// The one thing unit tests cannot see: that <c>sheets[].rows</c> (JSON of mixed types) and the
/// flat <c>operations[]</c> record bind through MCP SDK 1.2.0's stdio transport.
/// </summary>
public class ExcelWriteWorkflowTests
{
    [Fact]
    public async Task Create_then_update_via_stdio()
    {
        var created = TempPath(".xlsx");
        var updated = TempPath(".xlsx");
        try
        {
            await using var harness = await ServerHarness.StartAsync();

            var createResult = await harness.Client.CallToolAsync(
                "excel_create_workbook",
                new Dictionary<string, object?>
                {
                    ["path"] = created,
                    ["sheets"] = new object[]
                    {
                        new
                        {
                            name = "Data",
                            rows = new object?[][]
                            {
                                ["Name", "Amount", "Active", "Note", "When", "Double"],
                                ["Ada", 12.5, true, null, "2026-11-04T10:15", "=B2*2"],
                            },
                            columns = new object[] { new { column = "B", numberFormat = "0.00" } },
                        },
                        new { name = "Empty", header = false },
                    },
                });
            var createText = Text(createResult);
            Assert.Contains("\"cellsWritten\":12", createText);
            Assert.Contains("\"warnings\":[]", createText);
            Assert.Contains("\"name\":\"Empty\"", createText);

            var hashBefore = Sha256(created);
            var updateResult = await harness.Client.CallToolAsync(
                "excel_update_workbook",
                new Dictionary<string, object?>
                {
                    ["inputPath"] = created,
                    ["outputPath"] = updated,
                    ["operations"] = new object[]
                    {
                        new { op = "writeRange", sheet = "Data", startCell = "A3", rows = new object?[][] { ["Bob", 7, false, "x", "2026-12-31", "=B3*2"] } },
                        new { op = "formatRange", sheet = "Data", range = "A1:F1", fillColor = "#FFFF00" },
                        new { op = "freezePanes", sheet = "Data", cell = "B2" },
                    },
                });
            var updateText = Text(updateResult);
            Assert.Contains("\"operationsApplied\":3", updateText);
            Assert.Contains("\"cellsWritten\":6", updateText);
            Assert.Contains("\"warnings\":[]", updateText);
            Assert.Equal(hashBefore, Sha256(created));

            var read = await harness.Client.CallToolAsync(
                "excel_read_sheet",
                new Dictionary<string, object?> { ["path"] = updated, ["sheetName"] = "Data", ["range"] = "A1:F3", ["includeFormats"] = true });
            using var doc = JsonDocument.Parse(Text(read));
            var cells = doc.RootElement.GetProperty("cells").EnumerateArray()
                .ToDictionary(c => c.GetProperty("address").GetString()!, c => c);

            Assert.Equal("number", cells["B2"].GetProperty("valueType").GetString());
            Assert.Equal(12.5, cells["B2"].GetProperty("value").GetDouble());
            Assert.Equal("0.00", cells["B2"].GetProperty("numberFormat").GetString());
            Assert.Equal("boolean", cells["C2"].GetProperty("valueType").GetString());
            Assert.Equal("empty", cells["D2"].GetProperty("valueType").GetString());
            Assert.Equal("datetime", cells["E2"].GetProperty("valueType").GetString());
            Assert.Equal("2026-11-04T10:15:00", cells["E2"].GetProperty("value").GetString());
            Assert.Equal(25d, cells["F2"].GetProperty("value").GetDouble());        // create recalculated
            Assert.Equal("=B2*2", cells["F2"].GetProperty("formula").GetString());
            Assert.Equal("Bob", cells["A3"].GetProperty("value").GetString());
            Assert.Equal("datetime", cells["E3"].GetProperty("valueType").GetString());
            Assert.Equal("=B3*2", cells["F3"].GetProperty("formula").GetString());
            Assert.Equal(14d, cells["F3"].GetProperty("value").GetDouble());        // the written formula is evaluated; existing cached values are not
        }
        finally
        {
            foreach (var p in new[] { created, updated }) if (File.Exists(p)) File.Delete(p);
        }
    }

    [Fact]
    public async Task Invalid_operation_error_crosses_the_transport_with_its_code()
    {
        var created = TempPath(".xlsx");
        try
        {
            await using var harness = await ServerHarness.StartAsync();
            var result = await harness.Client.CallToolAsync(
                "excel_create_workbook",
                new Dictionary<string, object?>
                {
                    ["path"] = created,
                    ["sheets"] = new object[] { new { name = "Data", rows = new object?[][] { ["x"], ["=SUM(A1;A1)"] } } },
                });

            Assert.True(result.IsError);
            Assert.Contains("[invalid_operation] sheets[0]: cell A2", Text(result));
            Assert.False(File.Exists(created));
        }
        finally
        {
            if (File.Exists(created)) File.Delete(created);
        }
    }

    private static string Text(CallToolResult result) =>
        result.Content.OfType<TextContentBlock>().Single().Text;

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static string TempPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"mcpoffice-excel-write-integration-{Guid.NewGuid():N}{extension}");
}
