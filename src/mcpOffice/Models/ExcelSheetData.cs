namespace McpOffice.Models;

public sealed record ExcelSheetData(
    string Sheet,
    string Range,
    bool Truncated,
    IReadOnlyList<IReadOnlyList<object?>> Rows,
    IReadOnlyList<ExcelCellData> Cells);
