using DevExpress.Spreadsheet;
using McpOffice.Models;
using ModelContextProtocol;

namespace McpOffice.Services.Excel;

public sealed class ExcelWorkbookService : IExcelWorkbookService
{
    public IReadOnlyList<ExcelSheetInfo> ListSheets(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var workbook = LoadWorkbook(path);
            var sheets = new List<ExcelSheetInfo>();

            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                var worksheet = workbook.Worksheets[i];
                var usedRange = worksheet.GetUsedRange();
                var rowCount = usedRange.RowCount;
                var columnCount = usedRange.ColumnCount;

                sheets.Add(new ExcelSheetInfo(
                    i,
                    worksheet.Name,
                    worksheet.Visible,
                    "worksheet",
                    usedRange.GetReferenceA1(),
                    rowCount,
                    columnCount));
            }

            return sheets;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    private static Workbook LoadWorkbook(string path)
    {
        var workbook = new Workbook();
        workbook.LoadDocument(path);
        return workbook;
    }
}
