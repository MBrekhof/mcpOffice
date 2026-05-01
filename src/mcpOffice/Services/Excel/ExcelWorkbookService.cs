using DevExpress.Spreadsheet;
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;
using ModelContextProtocol;

namespace McpOffice.Services.Excel;

public sealed class ExcelWorkbookService : IExcelWorkbookService
{
    private const int DefaultSheetIndex = 0;

    public IReadOnlyList<ExcelSheetInfo> ListSheets(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var workbook = LoadWorkbook(path);
            var sheets = new List<ExcelSheetInfo>();

            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                var worksheet = workbook.Worksheets[i];
                var usedRange = worksheet.GetUsedRange();
                var rowCount = usedRange.RowCount;
                var columnCount = usedRange.ColumnCount;

                sheets.Add(new ExcelSheetInfo(
                    i,
                    worksheet.Name,
                    worksheet.Visible,
                    "worksheet",
                    usedRange.GetReferenceA1(),
                    rowCount,
                    columnCount));
            }

            return sheets;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public ExcelSheetData ReadSheet(
        string path,
        string? sheetName,
        int? sheetIndex,
        string? range,
        bool includeFormulas,
        bool includeFormats,
        int maxCells)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var workbook = LoadWorkbook(path);
            var worksheet = ResolveWorksheet(workbook, sheetName, sheetIndex);
            var cellRange = string.IsNullOrWhiteSpace(range)
                ? worksheet.GetUsedRange()
                : worksheet.Range[range];

            var rangeReference = cellRange.GetReferenceA1();
            var cellCount = checked(cellRange.RowCount * cellRange.ColumnCount);
            if (cellCount > maxCells)
            {
                throw ToolError.RangeTooLarge(rangeReference, cellCount, maxCells);
            }

            var rows = new List<IReadOnlyList<object?>>(cellRange.RowCount);
            var cells = new List<ExcelCellData>();

            for (var r = 0; r < cellRange.RowCount; r++)
            {
                var row = new List<object?>(cellRange.ColumnCount);
                for (var c = 0; c < cellRange.ColumnCount; c++)
                {
                    var cell = cellRange[r, c];
                    var value = GetCellValue(cell.Value);
                    row.Add(value);

                    cells.Add(new ExcelCellData(
                        cell.GetReferenceA1(),
                        value,
                        GetCellValueType(cell.Value),
                        includeFormulas && cell.HasFormula ? cell.Formula : null,
                        cell.DisplayText,
                        includeFormats ? cell.NumberFormat : null));
                }
                rows.Add(row);
            }

            return new ExcelSheetData(
                worksheet.Name,
                rangeReference,
                false,
                rows,
                cells);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public ExcelVbaProject ExtractVba(string path)
    {
        PathGuard.RequireExists(path);
        return new VbaProjectReader().Read(path);
    }

    private static Workbook LoadWorkbook(string path)
    {
        var workbook = new Workbook();
        workbook.LoadDocument(path);
        return workbook;
    }

    private static Worksheet ResolveWorksheet(Workbook workbook, string? sheetName, int? sheetIndex)
    {
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            var worksheet = workbook.Worksheets.FirstOrDefault(w =>
                string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheet is null)
            {
                throw ToolError.SheetNotFound(sheetName);
            }

            return worksheet;
        }

        var index = sheetIndex ?? DefaultSheetIndex;
        if (index < 0 || index >= workbook.Worksheets.Count)
        {
            throw ToolError.IndexOutOfRange(index, workbook.Worksheets.Count - 1);
        }

        return workbook.Worksheets[index];
    }

    private static object? GetCellValue(CellValue value)
    {
        if (value.IsEmpty)
        {
            return null;
        }

        if (value.IsBoolean)
        {
            return value.BooleanValue;
        }

        if (value.IsNumeric)
        {
            return value.NumericValue;
        }

        if (value.IsDateTime)
        {
            return value.DateTimeValue;
        }

        if (value.IsText)
        {
            return value.TextValue;
        }

        return value.ToString();
    }

    private static string GetCellValueType(CellValue value)
    {
        if (value.IsEmpty) return "empty";
        if (value.IsBoolean) return "boolean";
        if (value.IsNumeric) return "number";
        if (value.IsDateTime) return "datetime";
        if (value.IsText) return "text";
        return "unknown";
    }
}
