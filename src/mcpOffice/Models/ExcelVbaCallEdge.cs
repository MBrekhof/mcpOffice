namespace McpOffice.Models;

public sealed record ExcelVbaCallEdge(
    string From,
    string To,
    bool Resolved,
    ExcelVbaSiteRef Site);
