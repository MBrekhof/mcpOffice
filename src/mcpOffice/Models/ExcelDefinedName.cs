namespace McpOffice.Models;

public sealed record ExcelDefinedName(
    string Name,
    string? Scope,
    string RefersTo,
    string? Comment,
    bool IsHidden);
