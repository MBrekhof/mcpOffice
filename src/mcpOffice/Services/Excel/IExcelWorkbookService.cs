using McpOffice.Models;

namespace McpOffice.Services.Excel;

public interface IExcelWorkbookService
{
    IReadOnlyList<ExcelSheetInfo> ListSheets(string path);
    ExcelSheetData ReadSheet(
        string path,
        string? sheetName,
        int? sheetIndex,
        string? range,
        bool includeFormulas,
        bool includeFormats,
        int maxCells);
    ExcelVbaProject ExtractVba(string path);
    ExcelWorkbookMetadata GetMetadata(string path);
}
