using System.Globalization;
using DevExpress.Spreadsheet;
using SpreadsheetFormat = DevExpress.Spreadsheet.DocumentFormat;

namespace McpOffice.Tests.Excel;

internal static class TestExcelWorkbooks
{
    public static string Create(Action<Workbook> configure, SpreadsheetFormat? format = null)
    {
        var documentFormat = format ?? SpreadsheetFormat.Xlsx;
        var extension = documentFormat == SpreadsheetFormat.Xlsm ? ".xlsm" : ".xlsx";
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-excel-{Guid.NewGuid():N}{extension}");

        using var workbook = new Workbook();
        // DevExpress's formula parser (used by DefinedNames.Add, CellRange.Formula, etc.)
        // honors Workbook.Options.Culture for decimal separators and arg separators. On
        // non-English machines (e.g. nl-NL where "," is the decimal separator), test
        // fixtures using "=0.21" / "=1,2" would otherwise depend on the developer's locale.
        // Pin to invariant so test data stays portable.
        workbook.Options.Culture = CultureInfo.InvariantCulture;
        configure(workbook);
        workbook.SaveDocument(path, documentFormat);

        return path;
    }
}
