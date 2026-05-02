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

    [McpServerTool(Name = "excel_read_sheet")]
    [Description("Reads cell data from a worksheet or range. Returns rows plus addressed cell details. Uses maxCells to prevent accidental giant reads.")]
    public static object ExcelReadSheet(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path,
        [Description("Worksheet name. If omitted, sheetIndex is used.")] string? sheetName = null,
        [Description("0-based worksheet index used when sheetName is omitted. Defaults to 0.")] int? sheetIndex = null,
        [Description("Optional A1 range such as A1:D20. Defaults to the worksheet used range.")] string? range = null,
        [Description("Include formulas for formula cells.")] bool includeFormulas = true,
        [Description("Include number format strings.")] bool includeFormats = false,
        [Description("Maximum cells to return. Prevents accidental huge sheet reads.")] int maxCells = 50000)
        => Service.ReadSheet(path, sheetName, sheetIndex, range, includeFormulas, includeFormats, maxCells);

    [McpServerTool(Name = "excel_extract_vba")]
    [Description("Statically extracts VBA module source from an .xlsm workbook without launching Excel. Returns hasVbaProject and a list of {name, kind, lineCount, code}. For .xlsx or workbooks without macros, returns hasVbaProject=false and an empty list.")]
    public static object ExcelExtractVba(
        [Description("Absolute path to the .xlsm workbook")] string path)
        => Service.ExtractVba(path);

    [McpServerTool(Name = "excel_get_metadata")]
    [Description("Returns workbook document properties (author, title, subject, keywords, description, category, company, manager, application, lastModifiedBy, created, modified, printed) plus sheetCount.")]
    public static object ExcelGetMetadata(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path)
        => Service.GetMetadata(path);
}
