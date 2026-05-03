namespace McpOffice.Models;

public sealed record ExcelVbaObjectModelRef(
    string Module,
    string Procedure,
    int Line,
    string Api,
    string? Literal);
