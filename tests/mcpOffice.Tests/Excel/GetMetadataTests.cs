using McpOffice.Services.Excel;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel;

public class GetMetadataTests
{
    [Fact]
    public void GetMetadata_returns_document_properties_and_sheet_count()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Name = "Data";
            workbook.Worksheets.Add("Notes");

            var properties = workbook.DocumentProperties;
            properties.Author = "Martin";
            properties.Title = "Quarterly Numbers";
            properties.Subject = "Finance";
            properties.Keywords = "q1,finance";
            properties.Description = "Internal only";
            properties.Company = "Acme";
            properties.Category = "Reports";
        });

        try
        {
            var metadata = new ExcelWorkbookService().GetMetadata(path);

            Assert.Equal("Martin", metadata.Author);
            Assert.Equal("Quarterly Numbers", metadata.Title);
            Assert.Equal("Finance", metadata.Subject);
            Assert.Equal("q1,finance", metadata.Keywords);
            Assert.Equal("Internal only", metadata.Description);
            Assert.Equal("Acme", metadata.Company);
            Assert.Equal("Reports", metadata.Category);
            Assert.Equal(2, metadata.SheetCount);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void GetMetadata_throws_file_not_found_for_missing_workbook()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"mcpoffice-missing-{Guid.NewGuid():N}.xlsx");

        var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().GetMetadata(missing));

        Assert.Contains("file_not_found", ex.Message);
    }
}
