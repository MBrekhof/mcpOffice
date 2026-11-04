using System.Text.Json;
using DevExpress.Spreadsheet;
using McpOffice.Models;
using SpreadsheetFormat = DevExpress.Spreadsheet.DocumentFormat;

namespace McpOffice.Services.Excel.Write;

/// <summary>DevExpress-facing primitives shared by <c>CreateWorkbook</c> and <see cref="OperationApplier"/>.</summary>
internal static class WorkbookWriter
{
    public const int MaxAutoFitWidthInCharacters = 60;
    private const int MaxEvaluationPasses = 5;
    private static readonly System.Drawing.Color HeaderFill = ExcelWriteParsers.ParseHexColor("#D9D9D9");

    /// <summary>
    /// Writes <paramref name="rows"/> with its top-left at (1-based) <paramref name="startRow"/>/<paramref name="startColumn"/>,
    /// then evaluates the formulas it wrote (so forward references inside the block get a value) and
    /// reports any that evaluate to #NAME?: thrown when <paramref name="warnings"/> is null (create: no
    /// VBA can exist, so it is always a wrong function name), otherwise added as a warning (update: it
    /// may be a VBA user-defined function that only Excel can evaluate). Returns cells written.
    /// </summary>
    public static int WriteRows(Worksheet ws, int startRow, int startColumn, JsonElement[][] rows, List<string>? warnings)
    {
        var width = rows.Length == 0 ? 0 : rows.Max(r => r?.Length ?? 0);
        if (startRow - 1 + rows.Length > ExcelWriteParsers.MaxRows)
            throw new WriteValidationException($"{rows.Length} rows from row {startRow} run past the last row ({ExcelWriteParsers.MaxRows})");
        if (startColumn - 1 + width > ExcelWriteParsers.MaxColumns)
            throw new WriteValidationException($"{width} columns from column {ExcelWriteParsers.ColumnLetter(startColumn)} run past the last column (XFD)");

        var written = 0;
        var formulas = new List<Cell>();
        for (var r = 0; r < rows.Length; r++)
        {
            var row = rows[r];
            if (row is null) continue;
            for (var c = 0; c < row.Length; c++)
            {
                var cell = ws.Cells[startRow - 1 + r, startColumn - 1 + c];
                try
                {
                    if (Apply(cell, CellValueMapper.Map(row[c]))) formulas.Add(cell);
                }
                catch (WriteValidationException ex)
                {
                    throw new WriteValidationException($"cell {cell.GetReferenceA1()}: {ex.Message}");
                }
                written++;
            }
        }

        EvaluateWrittenFormulas(formulas);
        foreach (var cell in formulas)
        {
            if (!cell.Value.IsError || cell.Value.ErrorValue.Type != ErrorType.Name) continue;
            var message = $"cell {cell.GetReferenceA1()}: formula '{cell.Formula}' evaluates to #NAME? (unknown function). Use English function names (SUM, AVERAGE, IF)";
            if (warnings is null) throw new WriteValidationException(message);
            warnings.Add(message + "; a VBA user-defined function is fine and gets its value when Excel opens the file");
        }
        return written;
    }

    /// <returns>true when the cell now holds a formula.</returns>
    private static bool Apply(Cell cell, MappedCell mapped)
    {
        switch (mapped.Kind)
        {
            case MappedCellKind.Empty:
                cell.Value = CellValue.Empty;
                return false;
            case MappedCellKind.Number:
                cell.Value = (double)mapped.Value!;
                return false;
            case MappedCellKind.Boolean:
                cell.Value = (bool)mapped.Value!;
                return false;
            case MappedCellKind.Text:
                cell.Value = (string)mapped.Value!;
                return false;
            case MappedCellKind.DateTime:
                cell.Value = (DateTime)mapped.Value!;
                cell.NumberFormat = mapped.DefaultNumberFormat;
                return false;
            default:
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
                return true;
        }
    }

    /// <summary>
    /// Without this a written formula is saved with a cached value of 0, which a reader would trust.
    /// Only the written cells are evaluated, so pre-existing cached values stay (no CalculateFull on
    /// update: VBA UDFs and external links would become #NAME?). Repeated until stable so a written
    /// formula that refers to another written formula gets its value too.
    /// </summary>
    private static void EvaluateWrittenFormulas(List<Cell> formulas)
    {
        for (var pass = 0; pass < MaxEvaluationPasses && formulas.Count > 0; pass++)
        {
            var changed = false;
            foreach (var cell in formulas)
            {
                var before = cell.Value.ToString();
                cell.Calculate();
                if (cell.Value.ToString() != before) changed = true;
            }
            if (!changed) break;
        }
    }

    /// <summary>Rows, per-column formats, header defaults, auto-fit (capped) and then explicit widths.</summary>
    public static int WriteSheet(Worksheet ws, ExcelSheetSpec spec)
    {
        var rows = spec.Rows ?? [];
        var written = WriteRows(ws, 1, 1, rows, warnings: null);
        var rowCount = rows.Length;
        var columnCount = rows.Length == 0 ? 0 : rows.Max(r => r?.Length ?? 0);
        var firstDataRow = spec.Header ? 2 : 1;

        // Column formats first: auto-fit must measure the formatted text, not the raw number.
        var explicitWidths = new List<(int Column, double Width)>();
        var columns = spec.Columns ?? [];
        for (var i = 0; i < columns.Length; i++)
        {
            var column = columns[i] ?? throw new WriteValidationException($"columns[{i}] is null");
            var c = ExcelWriteParsers.ParseColumn(column.Column);
            var horizontal = column.HorizontalAlign is null ? (SpreadsheetHorizontalAlignment?)null : ParseHorizontal(column.HorizontalAlign);
            if (column.Width is double width)
            {
                if (width is < 0 or > 255) throw new WriteValidationException($"columns[{i}]: width {width} is outside 0..255 characters");
                explicitWidths.Add((c, width));
            }
            if (rowCount < firstDataRow) continue;
            var data = ws.Range.FromLTRB(c - 1, firstDataRow - 1, c - 1, rowCount - 1);
            if (column.NumberFormat is not null) data.NumberFormat = column.NumberFormat;
            if (horizontal is { } h) data.Alignment.Horizontal = h;
            if (column.WrapText is bool wrap) data.Alignment.WrapText = wrap;
        }

        if (columnCount > 0)
        {
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
            for (var c = 1; c <= columnCount; c++) ClampWidth(ws, c);
        }
        foreach (var (c, width) in explicitWidths) SetWidth(ws, c, width);

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
        Directory.CreateDirectory(dir);
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
