using System.Globalization;
using System.IO.Compression;
using DevExpress.Spreadsheet;
using McpOffice.Models;
using McpOffice.Services.Excel.Csv;
using McpOffice.Services.Excel.Vba;
using McpOffice.Services.Excel.Vba.Rendering;
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
            var worksheets = MaterializeWorksheets(workbook);

            for (var i = 0; i < worksheets.Count; i++)
            {
                var worksheet = worksheets[i];
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

    public ExcelExportCsvResult ExportCsv(
        string path,
        string outputPath,
        string? sheetName,
        int? sheetIndex,
        string? range,
        bool overwrite,
        int maxRows,
        bool trimTrailingEmptyRows = false)
    {
        PathGuard.RequireExists(path);
        PathGuard.RequireWritable(outputPath, overwrite);

        try
        {
            using var workbook = LoadWorkbook(path);
            var worksheet = ResolveWorksheet(workbook, sheetName, sheetIndex);
            var cellRange = string.IsNullOrWhiteSpace(range)
                ? worksheet.GetUsedRange()
                : worksheet.Range[range];

            var rangeReference = cellRange.GetReferenceA1();
            if (cellRange.RowCount > maxRows)
            {
                throw ToolError.RangeTooLargeRows(rangeReference, cellRange.RowCount, maxRows);
            }

            // Pandas-friendly trim: walk bottom-up, find the last row with at least one non-empty,
            // non-error cell. Real-world workbooks (e.g. ScreeningDB-V2.xlsm) often have used ranges
            // pinned to row N by formatting or trailing #REF! formulas, even though only rows 1..k
            // carry data. Trimming makes the CSV shape match the data instead of the grid.
            var effectiveRowCount = cellRange.RowCount;
            if (trimTrailingEmptyRows)
            {
                effectiveRowCount = ComputeLastNonEmptyRow(cellRange) + 1;
            }

            long bytesWritten;
            using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                using (var csv = new CsvWriter(fileStream))
                {
                    for (var r = 0; r < effectiveRowCount; r++)
                    {
                        var row = new object?[cellRange.ColumnCount];
                        for (var c = 0; c < cellRange.ColumnCount; c++)
                        {
                            row[c] = GetCellValue(cellRange[r, c].Value);
                        }
                        csv.WriteRow(row);
                    }
                }
                bytesWritten = fileStream.Length;
            }

            return new ExcelExportCsvResult(
                outputPath,
                effectiveRowCount,
                cellRange.ColumnCount,
                bytesWritten);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    private static int ComputeLastNonEmptyRow(CellRange cellRange)
    {
        var cols = cellRange.ColumnCount;
        for (var r = cellRange.RowCount - 1; r >= 0; r--)
        {
            for (var c = 0; c < cols; c++)
            {
                var v = cellRange[r, c].Value;
                // Considered "empty for trim purposes" — would emit an empty CSV field anyway:
                if (v.IsEmpty) continue;
                if (v.Type == CellValueType.Error) continue;
                // Real-world workbooks (e.g. ScreeningDB-V2.xlsm Compounds-N) extend the used
                // range with formulas like =IF(OR(...),"","value") that evaluate to "" — those
                // cells have IsEmpty=false, Type=Text, but produce an empty CSV field.
                if (v.IsText && string.IsNullOrEmpty(v.TextValue)) continue;
                return r;
            }
        }
        return -1;
    }

    public ExcelVbaProject ExtractVba(string path)
    {
        PathGuard.RequireExists(path);
        return new VbaProjectReader().Read(path);
    }

    public ExcelVbaAnalysis AnalyzeVba(
        string path,
        bool includeProcedures,
        bool includeCallGraph,
        bool includeReferences,
        string? moduleName = null)
    {
        PathGuard.RequireExists(path);

        try
        {
            var project = new VbaProjectReader().Read(path);
            return VbaSourceAnalyzer.Analyze(project, includeProcedures, includeCallGraph, includeReferences, moduleName);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public string RenderVbaCallgraph(
        string path,
        string format,
        string? moduleName,
        string? procedureName,
        int depth,
        string direction,
        string layout,
        int maxNodes)
    {
        PathGuard.RequireExists(path);

        ICallgraphRenderer renderer = format switch
        {
            "mermaid" => new MermaidCallgraphRenderer(),
            "dot" => new DotCallgraphRenderer(),
            _ => throw ToolError.InvalidRenderOption(
                $"format='{format}' is not one of mermaid, dot."),
        };

        if (layout != "clustered" && layout != "flat")
            throw ToolError.InvalidRenderOption(
                $"layout='{layout}' is not one of clustered, flat.");

        try
        {
            var project = new VbaProjectReader().Read(path);
            var analysis = VbaSourceAnalyzer.Analyze(
                project, includeProcedures: true, includeCallGraph: true, includeReferences: false);

            if (!analysis.HasVbaProject)
            {
                return renderer.Render(
                    new FilteredCallgraph(Array.Empty<CallgraphNode>(), Array.Empty<CallgraphEdge>()),
                    new CallgraphRenderOptions(layout));
            }

            var filtered = VbaCallgraphFilter.Apply(analysis,
                new CallgraphFilterOptions(
                    ModuleName: moduleName,
                    ProcedureName: procedureName,
                    Depth: depth,
                    Direction: direction,
                    MaxNodes: maxNodes));

            return renderer.Render(filtered, new CallgraphRenderOptions(layout));
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public ConversionHints SuggestVbaConversion(
        string path,
        string? moduleName,
        string? targetParadigm)
    {
        PathGuard.RequireExists(path);

        try
        {
            // Run the full analyzer with no module filter — the coupling scorer needs the whole graph.
            var project = new VbaProjectReader().Read(path);
            var analysis = VbaSourceAnalyzer.Analyze(
                project,
                includeProcedures: true,
                includeCallGraph: true,
                includeReferences: true,
                moduleName: null);

            return VbaConversionHintBuilder.Build(analysis, moduleName, targetParadigm);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public ExcelVbaEntryPointsResult ListVbaEntryPoints(string path, bool includeUnreachable, string? moduleName)
    {
        PathGuard.RequireExists(path);

        try
        {
            var project = new VbaProjectReader().Read(path);
            var sheets = new List<VbaEntryPointAnalyzer.SheetInput>();
            if (project.HasVbaProject)
            {
                // Drawing parts, VML and formulas come straight from the package: no DevExpress
                // workbook load (30 s on ScreeningDB-V2), and DevExpress has no macro-link API anyway.
                using var zip = ZipFile.OpenRead(path);
                foreach (var s in OpenXmlParts.ListSheets(zip))
                {
                    sheets.Add(new VbaEntryPointAnalyzer.SheetInput(
                        s.Name,
                        s.CodeName,
                        s.DrawingPartPath is null ? null : OpenXmlParts.ReadEntryText(zip, s.DrawingPartPath),
                        s.LegacyDrawingPartPath is null ? null : OpenXmlParts.ReadEntryText(zip, s.LegacyDrawingPartPath),
                        OpenXmlParts.ReadFormulas(zip, s.PartPath)));
                }
            }
            return VbaEntryPointAnalyzer.Analyze(path, project, sheets, includeUnreachable, moduleName);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public ExcelWorkbookMetadata GetMetadata(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var workbook = LoadWorkbook(path);
            var p = workbook.DocumentProperties;

            return new ExcelWorkbookMetadata(
                NullIfEmpty(p.Author),
                NullIfEmpty(p.Title),
                NullIfEmpty(p.Subject),
                NullIfEmpty(p.Keywords),
                NullIfEmpty(p.Description),
                NullIfEmpty(p.Category),
                NullIfEmpty(p.Company),
                NullIfEmpty(p.Manager),
                NullIfEmpty(p.Application),
                NullIfEmpty(p.LastModifiedBy),
                NormalizeDate(p.Created),
                NormalizeDate(p.Modified),
                NormalizeDate(p.Printed),
                workbook.Worksheets.Count);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public IReadOnlyList<ExcelDefinedName> ListDefinedNames(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var workbook = LoadWorkbook(path);
            var results = new List<ExcelDefinedName>();

            foreach (var name in workbook.DefinedNames)
            {
                results.Add(MapDefinedName(name, scope: null));
            }

            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var name in worksheet.DefinedNames)
                {
                    results.Add(MapDefinedName(name, scope: worksheet.Name));
                }
            }

            return results;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public IReadOnlyList<ExcelFormulaCell> ListFormulas(
        string path,
        string? sheetName,
        bool includeValues,
        int maxFormulas)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var workbook = LoadWorkbook(path);
            if (includeValues)
            {
                workbook.CalculateFull();
            }
            var targets = string.IsNullOrWhiteSpace(sheetName)
                ? workbook.Worksheets.AsEnumerable()
                : new[] { ResolveWorksheet(workbook, sheetName, sheetIndex: null) };

            var results = new List<ExcelFormulaCell>();
            foreach (var worksheet in targets)
            {
                var used = worksheet.GetUsedRange();
                if (used.RowCount == 0 || used.ColumnCount == 0)
                {
                    continue;
                }

                for (var r = 0; r < used.RowCount; r++)
                {
                    for (var c = 0; c < used.ColumnCount; c++)
                    {
                        var cell = used[r, c];
                        if (!cell.HasFormula)
                        {
                            continue;
                        }

                        if (results.Count >= maxFormulas)
                        {
                            throw ToolError.RangeTooLarge(used.GetReferenceA1(), results.Count + 1, maxFormulas);
                        }

                        results.Add(new ExcelFormulaCell(
                            worksheet.Name,
                            cell.GetReferenceA1(),
                            cell.Formula,
                            includeValues ? GetCellValue(cell.Value) : null,
                            includeValues ? GetCellValueType(cell.Value) : null));
                    }
                }
            }

            return results;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public ExcelWorkbookStructure GetStructure(
        string path,
        bool includeSheets,
        bool includeFormulaCounts,
        bool includeDefinedNames)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var workbook = LoadWorkbook(path);
            var worksheets = MaterializeWorksheets(workbook);

            var definedNameCount = workbook.DefinedNames.Count
                + worksheets.Sum(w => w.DefinedNames.Count);

            List<ExcelSheetStructure>? sheets = null;
            if (includeSheets)
            {
                sheets = new List<ExcelSheetStructure>(worksheets.Count);
                for (var i = 0; i < worksheets.Count; i++)
                {
                    var worksheet = worksheets[i];
                    var used = worksheet.GetUsedRange();
                    var formulaCount = includeFormulaCounts ? CountFormulas(used) : 0;

                    sheets.Add(new ExcelSheetStructure(
                        i,
                        worksheet.Name,
                        worksheet.Visible,
                        "worksheet",
                        used.GetReferenceA1(),
                        used.RowCount,
                        used.ColumnCount,
                        formulaCount,
                        worksheet.Tables.Count));
                }
            }

            List<ExcelDefinedName>? definedNames = null;
            if (includeDefinedNames)
            {
                definedNames = new List<ExcelDefinedName>(definedNameCount);
                foreach (var name in workbook.DefinedNames)
                {
                    definedNames.Add(MapDefinedName(name, scope: null));
                }
                foreach (var worksheet in worksheets)
                {
                    foreach (var name in worksheet.DefinedNames)
                    {
                        definedNames.Add(MapDefinedName(name, scope: worksheet.Name));
                    }
                }
            }

            return new ExcelWorkbookStructure(
                worksheets.Count,
                definedNameCount,
                sheets,
                definedNames);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    private static int CountFormulas(CellRange range)
    {
        if (range.RowCount == 0 || range.ColumnCount == 0)
        {
            return 0;
        }

        var count = 0;
        for (var r = 0; r < range.RowCount; r++)
        {
            for (var c = 0; c < range.ColumnCount; c++)
            {
                if (range[r, c].HasFormula) count++;
            }
        }
        return count;
    }

    private static ExcelDefinedName MapDefinedName(DefinedName name, string? scope) =>
        new(
            name.Name,
            scope,
            name.RefersTo ?? string.Empty,
            NullIfEmpty(name.Comment),
            name.Hidden);

    private static string? NullIfEmpty(string? value) =>
        string.IsNullOrEmpty(value) ? null : value;

    private static DateTime? NormalizeDate(DateTime value) =>
        value == default ? null : value;

    private static Workbook LoadWorkbook(string path)
    {
        var workbook = new Workbook();
        // Pin to InvariantCulture so formula text we return to the agent (DefinedName.RefersTo,
        // CellRange.Formula, etc.) uses "." as the decimal separator and "," as the argument
        // separator regardless of the host's locale. The MCP API contract is locale-neutral;
        // an agent on a Dutch dev box and a CI runner in en-US should see identical output.
        workbook.Options.Culture = CultureInfo.InvariantCulture;
        workbook.LoadDocument(path);
        return workbook;
    }

    // Workaround for a DevExpress.Spreadsheet bug observed on real-world workbooks
    // (e.g. RingOnderzoek.xlsm): NativeWorksheetCollection.get_Item throws
    // ArgumentOutOfRangeException at [0] even when Count >= 1, while foreach iteration
    // works fine. Materializing via enumeration sidesteps the broken indexer.
    private static List<Worksheet> MaterializeWorksheets(Workbook workbook)
    {
        var list = new List<Worksheet>(workbook.Worksheets.Count);
        foreach (var worksheet in workbook.Worksheets)
        {
            list.Add(worksheet);
        }
        return list;
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

        var worksheets = MaterializeWorksheets(workbook);
        var index = sheetIndex ?? DefaultSheetIndex;
        if (index < 0 || index >= worksheets.Count)
        {
            throw ToolError.IndexOutOfRange(index, worksheets.Count - 1);
        }

        return worksheets[index];
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

        // IsDateTime must be checked before IsNumeric: in DevExpress, date-formatted
        // cells report both true, but the caller-meaningful representation is DateTime,
        // not the Excel serial number.
        if (value.IsDateTime)
        {
            return value.DateTimeValue;
        }

        if (value.IsNumeric)
        {
            return value.NumericValue;
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
        if (value.IsDateTime) return "datetime";
        if (value.IsNumeric) return "number";
        if (value.IsText) return "text";
        return "unknown";
    }
}
