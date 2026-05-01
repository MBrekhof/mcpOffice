namespace McpOffice.Models;

public sealed record ExcelSheetInfo(
    int Index,
    string Name,
    bool Visible,
    string Kind,
    string? UsedRange,
    int RowCount,
    int ColumnCount);
