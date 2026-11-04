using System.IO.Compression;
using System.Text.Json;
using DevExpress.Spreadsheet;
using McpOffice.Models;
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Write;

internal static class WriteTestSupport
{
    public static string TempPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"mcpoffice-write-{Guid.NewGuid():N}{extension}");

    /// <summary>Rows as raw JSON, the way they arrive from the agent.</summary>
    public static JsonElement[][] Rows(params object?[][] rows) =>
        rows.Select(r => r.Select(c => JsonSerializer.SerializeToElement(c)).ToArray()).ToArray();

    public static ExcelWriteResult Create(string path, params ExcelSheetSpec[] sheets) =>
        new ExcelWorkbookService().CreateWorkbook(path, sheets, null, overwrite: false);

    public static ExcelWriteResult Update(string input, string output, params ExcelOperation[] ops) =>
        new ExcelWorkbookService().UpdateWorkbook(input, output, ops, overwrite: false);

    public static ExcelSheetData Read(string path, string? sheet = null, string? range = null) =>
        new ExcelWorkbookService().ReadSheet(path, sheet, null, range, includeFormulas: true, includeFormats: true, maxCells: 100_000);

    public static ExcelCellData Cell(this ExcelSheetData data, string address) => data.Cells.Single(c => c.Address == address);

    /// <summary>Loads the saved file with DevExpress for formatting assertions the read tools don't expose.</summary>
    public static Workbook Load(string path)
    {
        var wb = new Workbook();
        wb.Options.Culture = System.Globalization.CultureInfo.InvariantCulture;
        wb.LoadDocument(path);
        return wb;
    }

    /// <summary>Raw sheet XML, for view state (frozen panes) that Excel reads from the file.</summary>
    public static string SheetXml(string path, int sheetNumber = 1)
    {
        using var zip = ZipFile.OpenRead(path);
        using var reader = new StreamReader(zip.GetEntry($"xl/worksheets/sheet{sheetNumber}.xml")!.Open());
        return reader.ReadToEnd();
    }

    public static string WorkbookXml(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        using var reader = new StreamReader(zip.GetEntry("xl/workbook.xml")!.Open());
        return reader.ReadToEnd();
    }

    public static void Delete(params string[] paths)
    {
        foreach (var p in paths) if (File.Exists(p)) File.Delete(p);
    }
}
