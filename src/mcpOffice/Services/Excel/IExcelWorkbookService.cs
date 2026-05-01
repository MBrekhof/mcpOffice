using McpOffice.Models;

namespace McpOffice.Services.Excel;

public interface IExcelWorkbookService
{
    IReadOnlyList<ExcelSheetInfo> ListSheets(string path);
}
