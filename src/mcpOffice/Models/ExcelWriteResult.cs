namespace McpOffice.Models;

public sealed record ExcelWrittenSheet(string Name, string UsedRange);

public sealed record ExcelWriteResult(
    string OutputPath,
    IReadOnlyList<ExcelWrittenSheet> Sheets,
    int CellsWritten,
    int OperationsApplied,
    IReadOnlyList<string> Warnings);
