using System.ComponentModel;
using McpOffice.Services.Excel;
using ModelContextProtocol.Server;

namespace McpOffice.Tools;

[McpServerToolType]
public static class ExcelTools
{
    private static readonly IExcelWorkbookService Service = new ExcelWorkbookService();

    [McpServerTool(Name = "excel_list_sheets")]
    [Description("Returns worksheets in an Excel workbook with visibility and used-range summary.")]
    public static object ExcelListSheets(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path)
        => Service.ListSheets(path);

    [McpServerTool(Name = "excel_read_sheet")]
    [Description("Reads cell data from a worksheet or range. Returns rows plus addressed cell details. Uses maxCells to prevent accidental giant reads.")]
    public static object ExcelReadSheet(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path,
        [Description("Worksheet name. If omitted, sheetIndex is used.")] string? sheetName = null,
        [Description("0-based worksheet index used when sheetName is omitted. Defaults to 0.")] int? sheetIndex = null,
        [Description("Optional A1 range such as A1:D20. Defaults to the worksheet used range.")] string? range = null,
        [Description("Include formulas for formula cells.")] bool includeFormulas = true,
        [Description("Include number format strings.")] bool includeFormats = false,
        [Description("Maximum cells to return. Prevents accidental huge sheet reads.")] int maxCells = 50000)
        => Service.ReadSheet(path, sheetName, sheetIndex, range, includeFormulas, includeFormats, maxCells);

    [McpServerTool(Name = "excel_export_csv")]
    [Description("Streams a worksheet (or A1 range) to a CSV file on disk for pandas/polars consumption. RFC 4180 dialect, UTF-8 (no BOM), CRLF line endings, invariant-culture numbers, ISO 8601 datetimes, lowercase booleans. Formula cells emit their cached value (no formula text). Returns {outputPath, rowCount, columnCount, bytesWritten}.")]
    public static object ExcelExportCsv(
        [Description("Absolute path to the .xlsx/.xlsm input workbook")] string path,
        [Description("Absolute path to the .csv output file. Parent directory is created if missing.")] string outputPath,
        [Description("Worksheet name. If omitted, sheetIndex is used.")] string? sheetName = null,
        [Description("0-based worksheet index used when sheetName is omitted. Defaults to 0.")] int? sheetIndex = null,
        [Description("Optional A1 range such as A1:D20. Defaults to the worksheet used range.")] string? range = null,
        [Description("Overwrite outputPath if it already exists. Defaults to false.")] bool overwrite = false,
        [Description("Maximum rows to export. Defaults to 1,048,576 (Excel row ceiling). Trips range_too_large if exceeded.")] int maxRows = 1_048_576,
        [Description("When true, walks the resolved range bottom-up and truncates output at the last row that has any non-empty, non-error cell. Useful for workbooks whose used range is pinned far past the data by formatting or trailing #REF! formulas. Defaults to false (preserve the full range).")] bool trimTrailingEmptyRows = false)
        => Service.ExportCsv(path, outputPath, sheetName, sheetIndex, range, overwrite, maxRows, trimTrailingEmptyRows);

    [McpServerTool(Name = "excel_extract_vba")]
    [Description("Statically extracts VBA module source from an .xlsm workbook without launching Excel. Returns hasVbaProject and a list of {name, kind, lineCount, code}. For .xlsx or workbooks without macros, returns hasVbaProject=false and an empty list.")]
    public static object ExcelExtractVba(
        [Description("Absolute path to the .xlsm workbook")] string path)
        => Service.ExtractVba(path);

    [McpServerTool(Name = "excel_get_metadata")]
    [Description("Returns workbook document properties (author, title, subject, keywords, description, category, company, manager, application, lastModifiedBy, created, modified, printed) plus sheetCount.")]
    public static object ExcelGetMetadata(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path)
        => Service.GetMetadata(path);

    [McpServerTool(Name = "excel_list_defined_names")]
    [Description("Returns all defined names in the workbook. Each entry has {name, scope (null for workbook scope, sheet name for sheet scope), refersTo, comment, isHidden}.")]
    public static object ExcelListDefinedNames(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path)
        => Service.ListDefinedNames(path);

    [McpServerTool(Name = "excel_list_formulas")]
    [Description("Returns formula cells across the workbook (or a single sheet). Each entry has {sheet, address, formula, value?, valueType?}. When includeValues=true the workbook is recalculated and value/valueType are populated. maxFormulas caps the result; exceeding it raises range_too_large.")]
    public static object ExcelListFormulas(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path,
        [Description("Optional sheet name. When omitted, all sheets are scanned.")] string? sheetName = null,
        [Description("Recalculate and include cached values in each result.")] bool includeValues = false,
        [Description("Maximum number of formula cells to return.")] int maxFormulas = 10000)
        => Service.ListFormulas(path, sheetName, includeValues, maxFormulas);

    [McpServerTool(Name = "excel_get_structure")]
    [Description("Returns a workbook-level summary: sheetCount, definedNameCount, optional sheets array (per-sheet index/name/visibility/usedRange/row+columnCount/formulaCount/tableCount), and optional definedNames. Toggle the include* flags to keep payloads small on large workbooks.")]
    public static object ExcelGetStructure(
        [Description("Absolute path to the .xlsx/.xlsm workbook")] string path,
        [Description("Include the per-sheet array. Default true.")] bool includeSheets = true,
        [Description("Include formula counts per sheet (requires scanning each used range). Default true.")] bool includeFormulaCounts = true,
        [Description("Include the defined names array (workbook + sheet scoped). Default true.")] bool includeDefinedNames = true)
        => Service.GetStructure(path, includeSheets, includeFormulaCounts, includeDefinedNames);

    [McpServerTool(Name = "excel_analyze_vba")]
    [Description("Layers structural analysis on top of excel_extract_vba's source: procedures with signatures, event handlers, call graph, Excel object-model references (Worksheets/Range/Cells/...), and external dependencies (filesystem/database/network/automation/shell). Tiered output via toggles. Pass moduleName to scope the heavy arrays (modules / callGraph / references) to a single module on large workbooks; the summary stays whole-workbook. Returns hasVbaProject=false (with zeroed summary) for workbooks without a VBA project.")]
    public static object ExcelAnalyzeVba(
        [Description("Absolute path to the .xlsm/.xlsb workbook")] string path,
        [Description("Include the per-module procedure list. Default true.")] bool includeProcedures = true,
        [Description("Include the call graph edges. Default false (medium cost).")] bool includeCallGraph = false,
        [Description("Include object-model and dependency references. Default false (heaviest output).")] bool includeReferences = false,
        [Description("Optional case-insensitive VBA module name to scope the modules/callGraph/references arrays to. Summary remains whole-workbook. Throws module_not_found if the name is unknown.")] string? moduleName = null)
        => Service.AnalyzeVba(path, includeProcedures, includeCallGraph, includeReferences, moduleName);

    [McpServerTool(Name = "excel_render_vba_callgraph")]
    [Description("Renders the VBA call graph as Mermaid (default) or DOT for visual inspection. Layered on excel_analyze_vba. Use moduleName / procedureName / depth / direction to narrow on large workbooks; without filters, large workbooks throw graph_too_large. Returns the rendered string directly — no JSON wrapper.")]
    public static object ExcelRenderVbaCallgraph(
        [Description("Absolute path to the .xlsm/.xlsb workbook")] string path,
        [Description("Output format: 'mermaid' (default, renders inline in Markdown) or 'dot' (Graphviz).")] string format = "mermaid",
        [Description("Optional case-insensitive module name to scope the graph to a single module's neighbourhood.")] string? moduleName = null,
        [Description("Optional case-insensitive focal procedure name within moduleName. Requires moduleName.")] string? procedureName = null,
        [Description("BFS hops out from the focal procedure. Used only with procedureName. Default 2.")] int depth = 2,
        [Description("BFS direction: 'callees', 'callers', or 'both'. Used only with procedureName. Default 'both'.")] string direction = "both",
        [Description("Layout: 'clustered' (subgraph per module, default) or 'flat'.")] string layout = "clustered",
        [Description("Hard cap on rendered node count. Throws graph_too_large past this. Default 300.")] int maxNodes = 300)
        => Service.RenderVbaCallgraph(path, format, moduleName, procedureName, depth, direction, layout, maxNodes);

    [McpServerTool(Name = "excel_suggest_vba_conversion")]
    [Description("Conversion-hints layer over excel_analyze_vba. For each VBA procedure, emits multi-axis classification (trigger / purity / shape / dependencies), a human-readable rationale, and — when targetParadigm is set — a structured C# emission target (targetType / class / method / lifetime / blockers). Also returns workbook-wide module coupling: per-module Ca/Ce/instability + pairwise edge counts. moduleName scopes hints to a single module; coupling stays whole-workbook regardless. targetParadigm must be one of classLibrary, workerService, webApi, console.")]
    public static object ExcelSuggestVbaConversion(
        [Description("Absolute path to the .xlsm/.xlsb workbook")] string path,
        [Description("Optional case-insensitive VBA module name to scope per-procedure hints to. Coupling stays whole-workbook. Throws module_not_found if unknown.")] string? moduleName = null,
        [Description("Optional target paradigm: classLibrary | workerService | webApi | console. When set, every hint includes a structured csharpSuggestion. Throws unsupported_paradigm if the value is not in the supported set.")] string? targetParadigm = null)
        => Service.SuggestVbaConversion(path, moduleName, targetParadigm);

    [McpServerTool(Name = "excel_list_vba_entry_points")]
    [Description("What actually runs in a macro workbook, and what never can. Entry points: event handlers, Auto_* macros, macros wired to shapes (drawingN.xml) and form controls (vmlDrawingN.vml), Public Functions used as worksheet functions in cell formulas, and dynamic dispatch (Application.OnTime/OnKey/Run, .OnAction, CallByName with literal targets). Then walks the call graph from those entry points; unreachable[] lists procedures nothing can reach (confidence high|medium) — the migration scope cut. Kind vocabulary: eventHandler | autoMacro | shapeMacro | formControlMacro | worksheetFunction | dynamicDispatch. moduleName scopes both arrays; the summary stays whole-workbook. Returns hasVbaProject=false for workbooks without a VBA project.")]
    public static object ExcelListVbaEntryPoints(
        [Description("Absolute path to the .xlsm workbook")] string path,
        [Description("Include the unreachable[] array (reachability BFS). Default true.")] bool includeUnreachable = true,
        [Description("Optional case-insensitive VBA module name to scope entryPoints/unreachable to. Throws module_not_found if unknown.")] string? moduleName = null)
        => Service.ListVbaEntryPoints(path, includeUnreachable, moduleName);

    [McpServerTool(Name = "excel_map_vba_sheet_access")]
    [Description("The workbook's hidden data schema: per VBA procedure, which sheet and range / defined name it reads and writes. Resolves Worksheets(\"X\"), Sheets(n), sheet codenames (Blad1.Range), the sheet module's own unqualified Range/Cells, With blocks, one-assignment aliases (Set ws = …) and defined names (via refersTo). ActiveSheet and unqualified access outside a sheet module are reported as unresolved, never guessed. Records: {procedure, sheet{name,codeName}|null, target{kind: range|definedName|column|row|wholeSheet|dynamicCells, address?, definedName?}, mode: read|write|both, siteCount, unresolvedReason?}; sheets[] is the per-sheet rollup of readers/writers. On a big workbook call with includeRecords=false first (summary + sheets[] only, a few KB), then scope the records with moduleName / sheetName; maxRecords caps them and sets truncated=true. The summary stays whole-workbook.")]
    public static object ExcelMapVbaSheetAccess(
        [Description("Absolute path to the .xlsm workbook")] string path,
        [Description("Optional case-insensitive VBA module name to scope sheetAccess to. Throws module_not_found if unknown.")] string? moduleName = null,
        [Description("Optional sheet name to scope sheetAccess and sheets to. Throws sheet_not_found if unknown.")] string? sheetName = null,
        [Description("Include records whose sheet could not be resolved (activeSheet, aliasReassigned, unknownSheet, unknownName, dynamicSheet). Default true.")] bool includeUnresolved = true,
        [Description("Include the per-procedure sheetAccess[] records. Default true; pass false for the summary and per-sheet rollup only.")] bool includeRecords = true,
        [Description("Maximum number of sheetAccess records to return; truncated=true when cut. Default 300.")] int maxRecords = 300)
        => Service.MapVbaSheetAccess(path, moduleName, sheetName, includeUnresolved, includeRecords, maxRecords);

    [McpServerTool(Name = "excel_list_vba_form_controls")]
    [Description("The UI spec of each UserForm, inferred from its code-behind (the binary .frx designer part is not read): controls named by event handlers (cmdOK_Click), Me.<control> references, Hungarian-prefixed or VBE-default-named bare references (txt/cmd/lst/cbo/chk/opt/lbl/…, Label2/TextBox1/…) and 'As MSForms.<Type>' declarations. Each control has inferredType (MSForms type or Control) with typeConfidence declared | prefix | event | member | none, its events and referenced properties; formEvents lists the form's own handlers. formName scopes to one form; throws module_not_found if unknown.")]
    public static object ExcelListVbaFormControls(
        [Description("Absolute path to the .xlsm workbook")] string path,
        [Description("Optional case-insensitive UserForm module name (e.g. frmLogin).")] string? formName = null)
        => Service.ListVbaFormControls(path, formName);

    [McpServerTool(Name = "excel_compare_vba_corpus")]
    [Description("Finds VBA procedures shared across several .xlsm workbooks so they can be migrated once as a library instead of once per file. Pass exactly one of paths[] or directory (non-recursive *.xlsm). Tier identical = same normalised body (comments, whitespace and case ignored; name not part of the identity, so renamed copies still group); tier nearDuplicate = same name, body ≥ 90% line-similar. sharedModules[] lists module names present in several workbooks whose procedures are mostly shared. Per-workbook read failures land in workbooks[].error and the run continues. Loads every workbook: expect minutes on a directory of large files.")]
    public static object ExcelCompareVbaCorpus(
        [Description("Absolute paths to .xlsm workbooks (use this or directory).")] string[]? paths = null,
        [Description("Absolute directory whose *.xlsm files are compared, non-recursive (use this or paths).")] string? directory = null,
        [Description("Minimum number of distinct workbooks a procedure must appear in. Default 2.")] int minOccurrences = 2,
        [Description("Cap on sharedProcedures[] (sorted by occurrence count). Default 200; truncated=true when cut.")] int maxProcedures = 200,
        [Description("Also report same-named near-duplicate bodies (≥ 90% similar). Default true.")] bool includeNearDuplicates = true)
        => Service.CompareVbaCorpus(paths, directory, minOccurrences, maxProcedures, includeNearDuplicates);
}
