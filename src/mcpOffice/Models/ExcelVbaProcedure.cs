namespace McpOffice.Models;

public sealed record ExcelVbaProcedure(
    string Name,
    string FullyQualifiedName,
    string Kind,
    string? Scope,
    IReadOnlyList<ExcelVbaParameter> Parameters,
    string? ReturnType,
    int LineStart,
    int LineEnd,
    bool IsEventHandler,
    string? EventTarget);
