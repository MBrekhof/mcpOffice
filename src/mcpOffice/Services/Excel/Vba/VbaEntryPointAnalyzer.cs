using System.Text.RegularExpressions;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

/// <summary>
/// Composes v1's procedures + call graph with the workbook's drawing parts, formulas and the
/// dynamic-dispatch scanner into "what runs" (entry points) and "what never can" (unreachable).
/// Pure: the service hands it already-extracted inputs so tests can feed XML strings.
/// </summary>
internal static class VbaEntryPointAnalyzer
{
    public sealed record SheetInput(
        string Name,
        string? CodeName,
        string? DrawingXml,
        string? VmlDrawing,
        IReadOnlyList<(string Cell, string Formula)> Formulas);

    private static readonly StringComparer Ci = StringComparer.OrdinalIgnoreCase;
    private static readonly string[] AutoMacroNames = ["Auto_Open", "Auto_Close", "Auto_Activate", "Auto_Deactivate"];

    private sealed record ProcRef(ExcelVbaModuleAnalysis Module, ExcelVbaProcedure Proc);

    public static ExcelVbaEntryPointsResult Analyze(
        string path,
        ExcelVbaProject project,
        IReadOnlyList<SheetInput> sheets,
        bool includeUnreachable,
        string? moduleName,
        int maxItems = 500)
    {
        if (!project.HasVbaProject)
        {
            return new ExcelVbaEntryPointsResult(path, false,
                new ExcelVbaEntryPointsSummary(0, new Dictionary<string, int>(), 0, 0, 0, 0, 0, 0),
                [], includeUnreachable ? [] : null, false);
        }

        string? filter = null;
        if (!string.IsNullOrWhiteSpace(moduleName))
        {
            var match = project.Modules.FirstOrDefault(m => Ci.Equals(m.Name, moduleName))
                        ?? throw ToolError.ModuleNotFound(moduleName, project.Modules.Select(m => m.Name));
            filter = match.Name;
        }

        var analysis = VbaSourceAnalyzer.Analyze(project, includeProcedures: true, includeCallGraph: true, includeReferences: false);
        var modules = analysis.Modules ?? [];
        var procs = modules.SelectMany(m => m.Procedures.Select(p => new ProcRef(m, p))).ToList();
        // Property Get/Let/Set share an FQN; the first one stands for the group.
        var byFqn = procs.GroupBy(x => x.Proc.FullyQualifiedName, Ci).ToDictionary(g => g.Key, g => g.First(), Ci);
        var byName = procs.GroupBy(x => x.Proc.Name, Ci).ToDictionary(g => g.Key, g => g.ToList(), Ci);

        var entries = new List<ExcelVbaEntryPoint>();
        int unresolvedMacro = 0, dynamicUnresolved = 0, skippedParts = 0;
        bool callByNameSeen = false;

        foreach (var x in procs)
        {
            if (x.Proc.IsEventHandler)
                entries.Add(new ExcelVbaEntryPoint(x.Proc.FullyQualifiedName, "eventHandler", null, null, null, null, true, null));
            else if (x.Module.Kind == "standardModule" && AutoMacroNames.Contains(x.Proc.Name, Ci))
                entries.Add(new ExcelVbaEntryPoint(x.Proc.FullyQualifiedName, "autoMacro", null, null, null, null, true, null));
        }

        foreach (var sheet in sheets)
        {
            if (sheet.DrawingXml is not null)
            {
                try
                {
                    foreach (var sm in DrawingMacroExtractor.FromDrawingXml(sheet.DrawingXml))
                        entries.Add(ShapeEntry(sm, "shapeMacro", sheet.Name, byFqn, byName, ref unresolvedMacro));
                }
                catch (Exception) { skippedParts++; }
            }
            if (sheet.VmlDrawing is not null)
            {
                try
                {
                    foreach (var sm in DrawingMacroExtractor.FromVmlDrawing(sheet.VmlDrawing))
                        entries.Add(ShapeEntry(sm, "formControlMacro", sheet.Name, byFqn, byName, ref unresolvedMacro));
                }
                catch (Exception) { skippedParts++; }
            }
        }

        AddWorksheetFunctions(procs, sheets, entries);

        var extraEdges = new List<(string From, string To)>();
        foreach (var m in project.Modules)
        {
            if (string.IsNullOrEmpty(m.Code)) continue;
            var lines = VbaLineCleaner.Clean(m.Code);
            var scanned = VbaProcedureScanner.Scan(m.Kind, m.Name, lines);
            foreach (var dd in VbaDynamicDispatchScanner.Scan(m.Name, lines, scanned))
            {
                if (dd.Api == "CallByName") callByNameSeen = true;
                if (dd.TargetLiteral is null) { dynamicUnresolved++; continue; }

                var (targetModule, targetProc, parsable) = DrawingMacroExtractor.ParseMacroRef(dd.TargetLiteral);
                var resolved = parsable ? Resolve(targetModule, targetProc, byFqn, byName) : null;
                var site = new ExcelVbaSiteRef(dd.Module, dd.Procedure, dd.Line);
                entries.Add(new ExcelVbaEntryPoint(resolved?.Proc.FullyQualifiedName, "dynamicDispatch", null, null, null,
                    dd.TargetLiteral, resolved is not null, site));
                if (resolved is not null) extraEdges.Add(($"{dd.Module}.{dd.Procedure}", resolved.Proc.FullyQualifiedName));
                else unresolvedMacro++;
            }
        }

        var allFqns = byFqn.Keys.ToList();
        var entryFqns = entries.Where(e => e.Resolved && e.Procedure is not null).Select(e => e.Procedure!).Distinct(Ci);
        var reach = VbaCallGraphReachability.Compute(allFqns, entryFqns, analysis.CallGraph ?? [], extraEdges);

        var unreachable = reach.Unreachable.Select(fqn =>
        {
            var x = byFqn[fqn];
            var confident = Ci.Equals(x.Proc.Scope, "Private") && x.Module.Kind == "standardModule"
                            && dynamicUnresolved == 0 && !callByNameSeen;
            return new ExcelVbaUnreachableProcedure(fqn, x.Module.Name, x.Module.Kind, x.Proc.Scope,
                Math.Max(1, x.Proc.LineEnd - x.Proc.LineStart + 1), confident ? "high" : "medium");
        }).ToList();

        var byKind = new SortedDictionary<string, int>(StringComparer.Ordinal);
        foreach (var e in entries) byKind[e.Kind] = byKind.GetValueOrDefault(e.Kind) + 1;
        var summary = new ExcelVbaEntryPointsSummary(entries.Count, byKind, allFqns.Count, reach.Reachable.Count,
            unreachable.Count, unresolvedMacro, dynamicUnresolved, skippedParts);

        IEnumerable<ExcelVbaEntryPoint> ep = entries;
        IEnumerable<ExcelVbaUnreachableProcedure> un = unreachable;
        if (filter is not null)
        {
            ep = ep.Where(e => Ci.Equals(ModuleOf(e), filter));
            un = un.Where(u => Ci.Equals(u.Module, filter));
        }
        var epList = ep.OrderBy(e => e.Procedure ?? e.Target ?? "", Ci).ThenBy(e => e.Kind, StringComparer.Ordinal).ToList();
        var unList = un.OrderBy(u => u.Procedure, Ci).ToList();
        var truncated = epList.Count > maxItems || unList.Count > maxItems;

        return new ExcelVbaEntryPointsResult(path, true, summary,
            epList.Take(maxItems).ToList(),
            includeUnreachable ? unList.Take(maxItems).ToList() : null,
            truncated);
    }

    private static ExcelVbaEntryPoint ShapeEntry(
        DrawingMacroExtractor.ShapeMacro sm, string kind, string sheetName,
        Dictionary<string, ProcRef> byFqn, Dictionary<string, List<ProcRef>> byName, ref int unresolvedMacro)
    {
        var resolved = sm.Parsable ? Resolve(sm.TargetModule, sm.TargetProcedure, byFqn, byName) : null;
        if (resolved is null) unresolvedMacro++;
        return new ExcelVbaEntryPoint(resolved?.Proc.FullyQualifiedName, kind, sheetName, sm.ShapeName, null,
            sm.MacroRef, resolved is not null, null);
    }

    /// <summary>Module.Proc resolves exactly; a bare name resolves when unique, preferring standard modules.</summary>
    private static ProcRef? Resolve(string? module, string proc, Dictionary<string, ProcRef> byFqn, Dictionary<string, List<ProcRef>> byName)
    {
        if (module is not null)
            return byFqn.GetValueOrDefault($"{module}.{proc}");
        if (!byName.TryGetValue(proc, out var candidates)) return null;
        var standard = candidates.Where(c => c.Module.Kind == "standardModule").ToList();
        var pool = standard.Count > 0 ? standard : candidates;
        return pool.Count == 1 ? pool[0] : null;
    }

    private static void AddWorksheetFunctions(List<ProcRef> procs, IReadOnlyList<SheetInput> sheets, List<ExcelVbaEntryPoint> entries)
    {
        var functions = procs
            .Where(x => x.Module.Kind == "standardModule"
                        && x.Proc.Kind.Contains("function", StringComparison.OrdinalIgnoreCase)
                        && !Ci.Equals(x.Proc.Scope, "Private"))
            .GroupBy(x => x.Proc.Name, Ci)
            .ToDictionary(g => g.Key, g => g.First(), Ci);
        if (functions.Count == 0) return;

        // One alternation over every candidate name; a preceding '.' or word char rules out
        // sheet-qualified or longer identifiers (e.g. "Sheet1.Score(" or "MyScore(").
        var pattern = "(?<![\\w.])(?:" + string.Join("|", functions.Keys.Select(Regex.Escape)) + ")\\s*\\(";
        var regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        var cells = new Dictionary<string, List<string>>(Ci);
        foreach (var sheet in sheets)
            foreach (var (cell, formula) in sheet.Formulas)
                foreach (Match m in regex.Matches(formula))
                {
                    var name = m.Value.TrimEnd('(', ' ', '\t');
                    var fqn = functions[name].Proc.FullyQualifiedName;
                    if (!cells.TryGetValue(fqn, out var list)) cells[fqn] = list = [];
                    list.Add($"{sheet.Name}!{cell}");
                }

        foreach (var (fqn, list) in cells)
            entries.Add(new ExcelVbaEntryPoint(fqn, "worksheetFunction", null, null, list.Take(5).ToList(), null, true, null));
    }

    private static string? ModuleOf(ExcelVbaEntryPoint e)
    {
        if (e.Procedure is not null) return e.Procedure.Split('.', 2)[0];
        return e.Site?.Module;   // unresolved dynamic dispatch: attribute to the calling module; unresolved shapes have none
    }
}
