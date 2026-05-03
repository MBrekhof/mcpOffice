namespace McpOffice.Models;

public sealed record ExcelVbaParameter(
    string Name,
    string? Type,
    bool ByRef,
    bool Optional,
    string? DefaultValue);
