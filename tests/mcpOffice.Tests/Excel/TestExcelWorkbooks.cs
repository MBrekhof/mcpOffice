using DevExpress.Spreadsheet;
using SpreadsheetFormat = DevExpress.Spreadsheet.DocumentFormat;

namespace McpOffice.Tests.Excel;

internal static class TestExcelWorkbooks
{
    public static string Create(Action<Workbook> configure, SpreadsheetFormat? format = null)
    {
        var documentFormat = format ?? SpreadsheetFormat.Xlsx;
        var extension = documentFormat == SpreadsheetFormat.Xlsm ? ".xlsm" : ".xlsx";
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-excel-{Guid.NewGuid():N}{extension}");

        using var workbook = new Workbook();
        configure(workbook);
        workbook.SaveDocument(path, documentFormat);

        return path;
    }
}
