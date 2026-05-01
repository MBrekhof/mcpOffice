namespace McpOffice.Models;

public sealed record ExcelVbaModule(
    string Name,
    string Kind,
    int LineCount,
    string Code);
