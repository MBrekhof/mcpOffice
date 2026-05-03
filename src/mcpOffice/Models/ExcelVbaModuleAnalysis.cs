namespace McpOffice.Models;

public sealed record ExcelVbaModuleAnalysis(
    string Name,
    string Kind,
    bool Parsed,
    string? Reason,
    IReadOnlyList<ExcelVbaProcedure> Procedures);
