namespace McpOffice.Models;

public sealed record ExcelFormulaCell(
    string Sheet,
    string Address,
    string Formula,
    object? Value,
    string? ValueType);
