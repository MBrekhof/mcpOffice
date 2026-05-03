namespace McpOffice.Models;

public sealed record ExcelVbaDependency(
    string Module,
    string Procedure,
    int Line,
    string Kind,
    string? Target,
    string? Operation);
