using System.Text.Json;
using DevExpress.Spreadsheet;
using McpOffice.Models;
using SpreadsheetFormat = DevExpress.Spreadsheet.DocumentFormat;

namespace McpOffice.Services.Excel.Write;

/// <summary>DevExpress-facing primitives shared by <c>CreateWorkbook</c> and <see cref="OperationApplier"/>.</summary>
internal static class WorkbookWriter
{
    public const int MaxAutoFitWidthInCharacters = 60;
    private static readonly System.Drawing.Color HeaderFill = ExcelWriteParsers.ParseHexColor("#D9D9D9");

    /// <summary>Writes <paramref name="rows"/> with its top-left at (1-based) <paramref name="startRow"/>/<paramref name="startColumn"/>. Returns cells written.</summary>
    public static int WriteRows(Worksheet ws, int startRow, int startColumn, JsonElement[][] rows)
    {
        var written = 0;
        for (var r = 0; r < rows.Length; r++)
        {
            var row = rows[r];
            if (row is null) continue;
            for (var c = 0; c < row.Length; c++)
            {
                var cell = ws.Cells[startRow - 1 + r, startColumn - 1 + c];
                try
                {
                    Apply(cell, CellValueMapper.Map(row[c]));
                }
                catch (WriteValidationException ex)
                {
                    throw new WriteValidationException($"cell {cell.GetReferenceA1()}: {ex.Message}");
                }
                written++;
            }
        }
        return written;
    }

    private static void Apply(Cell cell, MappedCell mapped)
    {
        switch (mapped.Kind)
        {
            case MappedCellKind.Empty:
                cell.Value = CellValue.Empty;
                break;
            case MappedCellKind.Number:
                cell.Value = (double)mapped.Value!;
                break;
            case MappedCellKind.Boolean:
                cell.Value = (bool)mapped.Value!;
                break;
            case MappedCellKind.Text:
                cell.Value = (string)mapped.Value!;
                break;
            case MappedCellKind.DateTime:
                cell.Value = (DateTime)mapped.Value!;
                cell.NumberFormat = mapped.DefaultNumberFormat;
                break;
            case MappedCellKind.Formula:
                var formula = (string)mapped.Value!;
                try
                {
                    cell.Formula = formula;
                }
                catch (Exception ex)
                {
                    throw new WriteValidationException($"formula '{formula}' was rejected ({ex.Message.TrimEnd('.')}). Use English function names, ',' between arguments and '.' as the decimal separator.");
                }
                if (!cell.HasFormula)
                    throw new WriteValidationException($"formula '{formula}' was not accepted as a formula. Use English function names, ',' between arguments and '.' as the decimal separator.");
                break;
        }
    }

    /// <summary>Rows, then header defaults, then per-column settings, then auto-fit (capped) for columns without an explicit width.</summary>
    public static int WriteSheet(Worksheet ws, ExcelSheetSpec spec)
    {
        var rows = spec.Rows ?? [];
        var written = WriteRows(ws, 1, 1, rows);
        var rowCount = rows.Length;
        var columnCount = rows.Length == 0 ? 0 : rows.Max(r => r?.Length ?? 0);
        if (columnCount == 0) return written;

        if (spec.Header)
        {
            var header = ws.Range.FromLTRB(0, 0, columnCount - 1, 0);
            header.Font.Bold = true;
            header.Fill.BackgroundColor = HeaderFill;
            header.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin;
            ws.FreezeRows(0);
            ws.AutoFilter.Apply(ws.Range.FromLTRB(0, 0, columnCount - 1, rowCount - 1));
        }

        ws.Columns.AutoFit(0, columnCount - 1);
        var explicitWidth = new HashSet<int>();
        var firstDataRow = spec.Header ? 2 : 1;
        foreach (var column in spec.Columns ?? [])
        {
            var c = ExcelWriteParsers.ParseColumn(column.Column);
            if (column.Width is double width)
            {
                explicitWidth.Add(c);
                SetWidth(ws, c, width);
            }
            if (rowCount < firstDataRow) continue;
            var data = ws.Range.FromLTRB(c - 1, firstDataRow - 1, c - 1, rowCount - 1);
            if (column.NumberFormat is not null) data.NumberFormat = column.NumberFormat;
            if (column.HorizontalAlign is not null) data.Alignment.Horizontal = ParseHorizontal(column.HorizontalAlign);
            if (column.WrapText is bool wrap) data.Alignment.WrapText = wrap;
        }
        for (var c = 1; c <= columnCount; c++)
            if (!explicitWidth.Contains(c)) ClampWidth(ws, c);

        return written;
    }

    public static void SetWidth(Worksheet ws, int column, double widthInCharacters)
    {
        if (widthInCharacters is < 0 or > 255)
            throw new WriteValidationException($"width {widthInCharacters} is outside 0..255 characters");
        ws.Columns[column - 1].WidthInCharacters = widthInCharacters;
    }

    public static void ClampWidth(Worksheet ws, int column)
    {
        var col = ws.Columns[column - 1];
        if (col.WidthInCharacters > MaxAutoFitWidthInCharacters) col.WidthInCharacters = MaxAutoFitWidthInCharacters;
    }

    public static SpreadsheetHorizontalAlignment ParseHorizontal(string value) => value switch
    {
        "left" => SpreadsheetHorizontalAlignment.Left,
        "center" => SpreadsheetHorizontalAlignment.Center,
        "right" => SpreadsheetHorizontalAlignment.Right,
        _ => throw new WriteValidationException($"horizontalAlign '{value}' is not one of left, center, right"),
    };

    public static SpreadsheetVerticalAlignment ParseVertical(string value) => value switch
    {
        "top" => SpreadsheetVerticalAlignment.Top,
        "middle" => SpreadsheetVerticalAlignment.Center,
        "bottom" => SpreadsheetVerticalAlignment.Bottom,
        _ => throw new WriteValidationException($"verticalAlign '{value}' is not one of top, middle, bottom"),
    };

    public static SpreadsheetFormat FormatForPath(string path) =>
        path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ? SpreadsheetFormat.Xlsm : SpreadsheetFormat.Xlsx;

    /// <summary>Saves to a temp file beside the output and moves it into place, so a failure leaves no half-written file.</summary>
    public static void SaveAtomic(Workbook wb, string outputPath)
    {
        var dir = Path.GetDirectoryName(outputPath)!;
        var temp = Path.Combine(dir, $"~{Path.GetFileName(outputPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            wb.SaveDocument(temp, FormatForPath(outputPath));
            File.Move(temp, outputPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    public static ExcelWriteResult Summarize(Workbook wb, string outputPath, int cellsWritten, int operationsApplied, IReadOnlyList<string> warnings)
    {
        var sheets = new List<ExcelWrittenSheet>();
        foreach (var ws in wb.Worksheets)
            sheets.Add(new ExcelWrittenSheet(ws.Name, ws.GetUsedRange().GetReferenceA1()));
        return new ExcelWriteResult(outputPath, sheets, cellsWritten, operationsApplied, warnings);
    }
}
