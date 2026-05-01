namespace McpOffice.Models;

public sealed record ExcelCellData(
    string Address,
    object? Value,
    string ValueType,
    string? Formula,
    string DisplayText,
    string? NumberFormat);
