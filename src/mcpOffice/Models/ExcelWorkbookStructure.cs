namespace McpOffice.Models;

public sealed record ExcelWorkbookStructure(
    int SheetCount,
    int DefinedNameCount,
    IReadOnlyList<ExcelSheetStructure>? Sheets,
    IReadOnlyList<ExcelDefinedName>? DefinedNames);
