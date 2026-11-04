using System.Text.Json;

namespace McpOffice.Models;

/// <summary>One sheet of <c>excel_create_workbook</c>. Cell values are raw JSON: the JSON type decides the cell type.</summary>
public sealed record ExcelSheetSpec(
    string Name,
    JsonElement[][]? Rows = null,
    bool Header = true,
    ExcelColumnSpec[]? Columns = null);

public sealed record ExcelColumnSpec(
    string Column,
    string? NumberFormat = null,
    double? Width = null,
    string? HorizontalAlign = null,
    bool? WrapText = null);

/// <summary>
/// One operation of the closed vocabulary (design doc, "Operations"). Deliberately a flat record:
/// MCP SDK 1.2.0 input binding has no polymorphism to get wrong, and the agent sees one simple schema.
/// Which fields an <see cref="Op"/> needs is validated in <c>OperationApplier</c>.
/// </summary>
public sealed record ExcelOperation(
    string Op,
    string? Sheet = null,
    string? Name = null,
    int? Position = null,
    string? NewName = null,
    string? StartCell = null,
    JsonElement[][]? Rows = null,
    string? Range = null,
    string? What = null,
    bool? Bold = null,
    bool? Italic = null,
    bool? Underline = null,
    string? FontColor = null,
    double? FontSize = null,
    string? FillColor = null,
    string? NumberFormat = null,
    string? HorizontalAlign = null,
    string? VerticalAlign = null,
    bool? WrapText = null,
    string? Border = null,
    string? BorderScope = null,
    string? Columns = null,
    double? Width = null,
    bool? AutoFit = null,
    int? Row = null,
    string? Column = null,
    int? Count = null,
    string? Cell = null);
