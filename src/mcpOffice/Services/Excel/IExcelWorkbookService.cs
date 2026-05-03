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
    IReadOnlyList<ExcelDefinedName> ListDefinedNames(string path);
    IReadOnlyList<ExcelFormulaCell> ListFormulas(
        string path,
        string? sheetName,
        bool includeValues,
        int maxFormulas);
    ExcelWorkbookStructure GetStructure(
        string path,
        bool includeSheets,
        bool includeFormulaCounts,
        bool includeDefinedNames);
    ExcelVbaAnalysis AnalyzeVba(
        string path,
        bool includeProcedures,
        bool includeCallGraph,
        bool includeReferences,
        string? moduleName = null);
}
