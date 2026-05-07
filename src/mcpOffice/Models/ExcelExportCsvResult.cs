// src/mcpOffice/Models/ExcelExportCsvResult.cs
namespace McpOffice.Models;

public sealed record ExcelExportCsvResult(
    string OutputPath,
    int RowCount,
    int ColumnCount,
    long BytesWritten);
