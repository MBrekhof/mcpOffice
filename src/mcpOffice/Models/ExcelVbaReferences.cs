namespace McpOffice.Models;

public sealed record ExcelVbaReferences(
    IReadOnlyList<ExcelVbaObjectModelRef> ObjectModel,
    IReadOnlyList<ExcelVbaDependency> Dependencies);
