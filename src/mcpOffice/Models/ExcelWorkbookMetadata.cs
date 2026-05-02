namespace McpOffice.Models;

public sealed record ExcelWorkbookMetadata(
    string? Author,
    string? Title,
    string? Subject,
    string? Keywords,
    string? Description,
    string? Category,
    string? Company,
    string? Manager,
    string? Application,
    string? LastModifiedBy,
    DateTime? Created,
    DateTime? Modified,
    DateTime? Printed,
    int SheetCount);
