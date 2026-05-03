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

    [McpServerTool(Name = "excel_list_defined_names")]
    [Description("Returns all defined names in the workbook. Each entry has {name, scope (null for workbook scope, sheet name for sheet scope), refersTo, comment, isHidden}.")]
    public static object ExcelListDefinedNames(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path)
        => Service.ListDefinedNames(path);

    [McpServerTool(Name = "excel_list_formulas")]
    [Description("Returns formula cells across the workbook (or a single sheet). Each entry has {sheet, address, formula, value?, valueType?}. When includeValues=true the workbook is recalculated and value/valueType are populated. maxFormulas caps the result; exceeding it raises range_too_large.")]
    public static object ExcelListFormulas(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path,
        [Description("Optional sheet name. When omitted, all sheets are scanned.")] string? sheetName = null,
        [Description("Recalculate and include cached values in each result.")] bool includeValues = false,
        [Description("Maximum number of formula cells to return.")] int maxFormulas = 10000)
        => Service.ListFormulas(path, sheetName, includeValues, maxFormulas);

    [McpServerTool(Name = "excel_get_structure")]
    [Description("Returns a workbook-level summary: sheetCount, definedNameCount, optional sheets array (per-sheet index/name/visibility/usedRange/row+columnCount/formulaCount/tableCount), and optional definedNames. Toggle the include* flags to keep payloads small on large workbooks.")]
    public static object ExcelGetStructure(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path,
        [Description("Include the per-sheet array. Default true.")] bool includeSheets = true,
        [Description("Include formula counts per sheet (requires scanning each used range). Default true.")] bool includeFormulaCounts = true,
        [Description("Include the defined names array (workbook + sheet scoped). Default true.")] bool includeDefinedNames = true)
        => Service.GetStructure(path, includeSheets, includeFormulaCounts, includeDefinedNames);

    [McpServerTool(Name = "excel_analyze_vba")]
    [Description("Layers structural analysis on top of excel_extract_vba's source: procedures with signatures, event handlers, call graph, Excel object-model references (Worksheets/Range/Cells/...), and external dependencies (file/database/network/automation/shell). Tiered output via toggles. Returns hasVbaProject=false (with zeroed summary) for workbooks without a VBA project.")]
    public static object ExcelAnalyzeVba(
        [Description("Absolute path to the .xlsm/.xlsb workbook")] string path,
        [Description("Include the per-module procedure list. Default true.")] bool includeProcedures = true,
        [Description("Include the call graph edges. Default false (medium cost).")] bool includeCallGraph = false,
        [Description("Include object-model and dependency references. Default false (heaviest output).")] bool includeReferences = false)
        => Service.AnalyzeVba(path, includeProcedures, includeCallGraph, includeReferences);
}
