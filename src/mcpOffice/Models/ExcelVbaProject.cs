namespace McpOffice.Models;

public sealed record ExcelVbaProject(
    bool HasVbaProject,
    IReadOnlyList<ExcelVbaModule> Modules);
