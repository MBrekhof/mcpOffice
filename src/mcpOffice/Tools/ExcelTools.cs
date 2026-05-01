using System.ComponentModel;
using McpOffice.Services.Excel;
using ModelContextProtocol.Server;

namespace McpOffice.Tools;

[McpServerToolType]
public static class ExcelTools
{
    private static readonly IExcelWorkbookService Service = new ExcelWorkbookService();

    [McpServerTool(Name = "excel_list_sheets")]
    [Description("Returns worksheets in an Excel workbook with visibility and used-range summary.")]
    public static object ExcelListSheets(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path)
        => Service.ListSheets(path);
}
