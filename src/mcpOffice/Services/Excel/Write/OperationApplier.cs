using System.Drawing;
using DevExpress.Spreadsheet;
using McpOffice.Models;
using ModelContextProtocol;

namespace McpOffice.Services.Excel.Write;

/// <summary>
/// Applies the closed operation vocabulary to an in-memory workbook, in order. The first failure
/// throws with the operation's index and <c>op</c>; the caller saves nothing (atomicity).
/// </summary>
internal static class OperationApplier
{
    public static readonly string[] ValidOps =
    [
        "addSheet", "renameSheet", "deleteSheet", "writeRange", "clearRange", "formatRange",
        "mergeCells", "unmergeCells", "setColumnWidth", "insertRows", "deleteRows",
        "insertColumns", "deleteColumns", "freezePanes", "autoFilter",
    ];

    /// <param name="warnings">Collects #NAME? results of written formulas; null makes them errors (create).</param>
    /// <returns>Cells written by <c>writeRange</c> operations.</returns>
    public static int Apply(Workbook wb, IReadOnlyList<ExcelOperation> ops, List<string>? warnings)
    {
        var cellsWritten = 0;
        for (var i = 0; i < ops.Count; i++)
        {
            var op = ops[i];
            if (op is null) throw ToolError.InvalidOperation($"operations[{i}]", "operation is null");
            var where = $"operations[{i}] ({op.Op})";
            try
            {
                cellsWritten += ApplyOne(wb, op, where, warnings);
            }
            catch (WriteValidationException ex)
            {
                throw ToolError.InvalidOperation(where, ex.Message);
            }
            catch (Exception ex) when (ex is not McpException)
            {
                // DevExpress refused the operation (cells shifted off the sheet, a multi-area range, ...).
                throw ToolError.InvalidOperation(where, ex.Message.TrimEnd('.'));
            }
        }
        return cellsWritten;
    }

    private static int ApplyOne(Workbook wb, ExcelOperation op, string where, List<string>? warnings)
    {
        switch (op.Op)
        {
            case "addSheet":
            {
                var name = Require(op.Name, "name");
                ExcelWriteParsers.ValidateSheetName(name);
                RequireNoSheet(wb, name, except: null);
                if (op.Position is int position)
                {
                    if (position < 0 || position > wb.Worksheets.Count)
                        throw new WriteValidationException($"position {position} is outside 0..{wb.Worksheets.Count}");
                    wb.Worksheets.Insert(position, name);
                }
                else wb.Worksheets.Add(name);
                return 0;
            }
            case "renameSheet":
            {
                var ws = Sheet(wb, op, where);
                var newName = Require(op.NewName, "newName");
                ExcelWriteParsers.ValidateSheetName(newName);
                RequireNoSheet(wb, newName, except: ws);
                ws.Name = newName;
                return 0;
            }
            case "deleteSheet":
            {
                var ws = Sheet(wb, op, where);
                if (wb.Worksheets.Count == 1) throw new WriteValidationException("cannot delete the last sheet of a workbook");
                if (ws.Visible && !wb.Worksheets.Any(w => w.Visible && !ReferenceEquals(w, ws)))
                    throw new WriteValidationException($"cannot delete '{ws.Name}': it is the only visible sheet (the others are hidden)");
                wb.Worksheets.Remove(ws);
                return 0;
            }
            case "writeRange":
            {
                var ws = Sheet(wb, op, where);
                var (row, column) = ExcelWriteParsers.ParseCell(Require(op.StartCell, "startCell"));
                return WorkbookWriter.WriteRows(ws, row, column, op.Rows ?? throw new WriteValidationException("rows is required"), warnings);
            }
            case "clearRange":
            {
                var ws = Sheet(wb, op, where);
                var range = Range(ws, op.Range);
                switch (Require(op.What, "what"))
                {
                    case "contents": ws.ClearContents(range); break;
                    case "formats": ws.ClearFormats(range); break;
                    case "all": ws.Clear(range); break;
                    default: throw new WriteValidationException($"what '{op.What}' is not one of contents, formats, all");
                }
                return 0;
            }
            case "formatRange":
                FormatRange(Sheet(wb, op, where), op);
                return 0;
            case "mergeCells":
            {
                var ws = Sheet(wb, op, where);
                ws.MergeCells(Range(ws, op.Range));
                return 0;
            }
            case "unmergeCells":
            {
                var ws = Sheet(wb, op, where);
                ws.UnMergeCells(Range(ws, op.Range));
                return 0;
            }
            case "setColumnWidth":
            {
                var ws = Sheet(wb, op, where);
                var (first, last) = ExcelWriteParsers.ParseColumnSpan(Require(op.Columns, "columns"));
                var autoFit = op.AutoFit == true;
                if (op.Width.HasValue == autoFit)
                    throw new WriteValidationException("give exactly one of width or autoFit: true");
                if (autoFit)
                {
                    ws.Columns.AutoFit(first - 1, last - 1);
                    for (var c = first; c <= last; c++) WorkbookWriter.ClampWidth(ws, c);
                }
                else
                {
                    for (var c = first; c <= last; c++) WorkbookWriter.SetWidth(ws, c, op.Width!.Value);
                }
                return 0;
            }
            case "insertRows":
            {
                var ws = Sheet(wb, op, where);
                var (row, count) = RowSpan(op, insert: true);
                ws.Rows.Insert(row - 1, count);
                return 0;
            }
            case "deleteRows":
            {
                var ws = Sheet(wb, op, where);
                var (row, count) = RowSpan(op, insert: false);
                ws.Rows.Remove(row - 1, count);
                return 0;
            }
            case "insertColumns":
            {
                var ws = Sheet(wb, op, where);
                var (column, count) = ColumnSpan(op, insert: true);
                ws.Columns.Insert(column - 1, count);
                return 0;
            }
            case "deleteColumns":
            {
                var ws = Sheet(wb, op, where);
                var (column, count) = ColumnSpan(op, insert: false);
                ws.Columns.Remove(column - 1, count);
                return 0;
            }
            case "freezePanes":
            {
                var ws = Sheet(wb, op, where);
                if (op.Cell is null) { ws.UnfreezePanes(); return 0; }
                var (row, column) = ExcelWriteParsers.ParseCell(op.Cell);
                var rows = row - 1;      // rows above the cell stay frozen
                var columns = column - 1;
                if (rows == 0 && columns == 0) throw new WriteValidationException("cell A1 freezes nothing; pass A2 for row 1, B2 for row 1 and column A, or null to unfreeze");
                if (rows > 0 && columns > 0) ws.FreezePanes(rows - 1, columns - 1);
                else if (rows > 0) ws.FreezeRows(rows - 1);
                else ws.FreezeColumns(columns - 1);
                return 0;
            }
            case "autoFilter":
            {
                var ws = Sheet(wb, op, where);
                if (op.Range is null) ws.AutoFilter.Disable();
                else ws.AutoFilter.Apply(Range(ws, op.Range));
                return 0;
            }
            default:
                throw new WriteValidationException($"unknown op '{op.Op}'. Valid ops: {string.Join(", ", ValidOps)}");
        }
    }

    private static void FormatRange(Worksheet ws, ExcelOperation op)
    {
        var range = Range(ws, op.Range);
        var touched = false;
        if (op.Bold is bool bold) { range.Font.Bold = bold; touched = true; }
        if (op.Italic is bool italic) { range.Font.Italic = italic; touched = true; }
        if (op.Underline is bool underline) { range.Font.UnderlineType = underline ? UnderlineType.Single : UnderlineType.None; touched = true; }
        if (op.FontColor is not null) { range.Font.Color = ExcelWriteParsers.ParseHexColor(op.FontColor); touched = true; }
        if (op.FontSize is double size)
        {
            if (size is < 1 or > 409) throw new WriteValidationException(FormattableString.Invariant($"fontSize {size} is outside 1..409"));
            range.Font.Size = size;
            touched = true;
        }
        if (op.FillColor is not null) { range.Fill.BackgroundColor = ExcelWriteParsers.ParseHexColor(op.FillColor); touched = true; }
        if (op.NumberFormat is not null) { range.NumberFormat = op.NumberFormat; touched = true; }
        if (op.HorizontalAlign is not null) { range.Alignment.Horizontal = WorkbookWriter.ParseHorizontal(op.HorizontalAlign); touched = true; }
        if (op.VerticalAlign is not null) { range.Alignment.Vertical = WorkbookWriter.ParseVertical(op.VerticalAlign); touched = true; }
        if (op.WrapText is bool wrap) { range.Alignment.WrapText = wrap; touched = true; }
        if (op.Border is not null)
        {
            var style = op.Border switch
            {
                "none" => BorderLineStyle.None,
                "thin" => BorderLineStyle.Thin,
                "medium" => BorderLineStyle.Medium,
                "thick" => BorderLineStyle.Thick,
                _ => throw new WriteValidationException($"border '{op.Border}' is not one of none, thin, medium, thick"),
            };
            switch (op.BorderScope ?? "all")
            {
                case "all": range.Borders.SetAllBorders(Color.Black, style); break;
                case "outline": range.Borders.SetOutsideBorders(Color.Black, style); break;
                default: throw new WriteValidationException($"borderScope '{op.BorderScope}' is not one of outline, all");
            }
            touched = true;
        }
        else if (op.BorderScope is not null)
        {
            throw new WriteValidationException("borderScope needs border");
        }
        if (!touched) throw new WriteValidationException("formatRange needs at least one of bold, italic, underline, fontColor, fontSize, fillColor, numberFormat, horizontalAlign, verticalAlign, wrapText, border");
    }

    private static Worksheet Sheet(Workbook wb, ExcelOperation op, string where)
    {
        var name = Require(op.Sheet, "sheet");
        return Find(wb, name) ?? throw ToolError.SheetNotFound($"{name} ({where})");
    }

    private static Worksheet? Find(Workbook wb, string name)
    {
        foreach (var ws in wb.Worksheets)
            if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase)) return ws;
        return null;
    }

    private static void RequireNoSheet(Workbook wb, string name, Worksheet? except)
    {
        var existing = Find(wb, name);
        if (existing is not null && !ReferenceEquals(existing, except))
            throw new WriteValidationException($"a sheet named '{name}' already exists (names are case-insensitive)");
    }

    private static CellRange Range(Worksheet ws, string? reference)
    {
        var r = Require(reference, "range");
        CellRange range;
        try
        {
            range = ws.Range[r];
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw new WriteValidationException($"range '{r}' is not a valid A1 reference: {ex.Message}");
        }
        if (!ReferenceEquals(range.Worksheet, ws) && !string.Equals(range.Worksheet?.Name, ws.Name, StringComparison.OrdinalIgnoreCase))
            throw new WriteValidationException($"range '{r}' refers to another sheet; use 'sheet' to pick the sheet and a plain A1 range");
        return range;
    }

    // An insert shifts everything from the index on; inserting Max - index + 1 or more would shift the
    // whole remainder off the sheet (DevExpress then silently does nothing), hence the stricter bound.
    private static (int Row, int Count) RowSpan(ExcelOperation op, bool insert)
    {
        var row = op.Row ?? throw new WriteValidationException("row is required");
        if (row < 1 || row > ExcelWriteParsers.MaxRows) throw new WriteValidationException($"row {row} is outside 1..{ExcelWriteParsers.MaxRows}");
        var count = Count(op);
        if ((long)row - 1 + count > ExcelWriteParsers.MaxRows - (insert ? 1 : 0))
            throw new WriteValidationException($"row {row} + count {count} runs past the last row ({ExcelWriteParsers.MaxRows})");
        return (row, count);
    }

    private static (int Column, int Count) ColumnSpan(ExcelOperation op, bool insert)
    {
        var column = ExcelWriteParsers.ParseColumn(Require(op.Column, "column"));
        var count = Count(op);
        if ((long)column - 1 + count > ExcelWriteParsers.MaxColumns - (insert ? 1 : 0))
            throw new WriteValidationException($"column {op.Column} + count {count} runs past the last column (XFD, {ExcelWriteParsers.MaxColumns})");
        return (column, count);
    }

    private static int Count(ExcelOperation op)
    {
        var count = op.Count ?? 1;
        if (count < 1) throw new WriteValidationException($"count {count} must be at least 1");
        return count;
    }

    private static string Require(string? value, string field) =>
        string.IsNullOrWhiteSpace(value) ? throw new WriteValidationException($"{field} is required") : value;
}
