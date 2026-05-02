namespace McpOffice.Models;

public sealed record ExcelSheetStructure(
    int Index,
    string Name,
    bool Visible,
    string Kind,
    string? UsedRange,
    int RowCount,
    int ColumnCount,
    int FormulaCount,
    int TableCount);
