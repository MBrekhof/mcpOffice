using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class InsertTableTests
{
    [Fact]
    public void InsertTable_writes_header_row_and_body_rows()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-table-{Guid.NewGuid():N}.docx");
        try
        {
            var service = new WordDocumentService();
            service.CreateBlank(path, overwrite: false);

            service.InsertTable(
                path,
                atIndex: 0,
                headers: new[] { "Name", "Age" },
                rows: new[] { new[] { "Alice", "30" }, new[] { "Bob", "25" } });

            var structured = service.ReadStructured(path);
            var table = Assert.Single(structured.Tables);
            Assert.Equal(3, table.Rows.Count);
            Assert.Equal("Name", table.Rows[0][0]);
            Assert.Equal("Age", table.Rows[0][1]);
            Assert.Equal("Alice", table.Rows[1][0]);
            Assert.Equal("25", table.Rows[2][1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
