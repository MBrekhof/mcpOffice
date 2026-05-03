namespace McpOffice.Models;

public sealed record ExcelVbaSiteRef(
    string Module,
    string Procedure,
    int Line);
