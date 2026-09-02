using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

/// <summary>
/// Runs <see cref="VbaSheetAccessResolver"/> over every module and aggregates the sites into
/// (procedure, sheet, target, mode) records plus a per-sheet rollup. Pure.
/// </summary>
internal static class VbaSheetAccessAnalyzer
{
    private static readonly StringComparer Ci = StringComparer.OrdinalIgnoreCase;

    public static ExcelVbaSheetAccessResult Analyze(
        string path,
        ExcelVbaProject project,
        IReadOnlyList<VbaSheetAccessResolver.SheetName> sheets,
        IReadOnlyList<VbaSheetAccessResolver.DefinedName> definedNames,
        string? moduleName,
        string? sheetName,
        bool includeUnresolved,
        bool includeRecords = true,
        int maxRecords = 100)   // 300 measured 59 KB on Air.xlsm, still over Claude Code's tool-result cap
    {
        if (!project.HasVbaProject)
            return new ExcelVbaSheetAccessResult(path, false, new ExcelVbaSheetAccessSummary(0, 0, 0, 0, 0), [], [], false);

        string? moduleFilter = null;
        if (!string.IsNullOrWhiteSpace(moduleName))
        {
            var m = project.Modules.FirstOrDefault(x => Ci.Equals(x.Name, moduleName))
                    ?? throw ToolError.ModuleNotFound(moduleName, project.Modules.Select(x => x.Name));
            moduleFilter = m.Name;
        }
        string? sheetFilter = null;
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            var s = sheets.FirstOrDefault(x => Ci.Equals(x.Name, sheetName))
                    ?? throw ToolError.SheetNotFound(sheetName);
            sheetFilter = s.Name;
        }

        var sites = new List<VbaSheetAccessResolver.AccessSite>();
        foreach (var m in project.Modules)
        {
            if (string.IsNullOrEmpty(m.Code)) continue;
            var lines = VbaLineCleaner.Clean(m.Code);
            var procs = VbaProcedureScanner.Scan(m.Kind, m.Name, lines);
            sites.AddRange(VbaSheetAccessResolver.Resolve(m.Name, m.Kind, lines, procs, sheets, definedNames));
        }

        var codeNameBySheet = sheets.ToDictionary(s => s.Name, s => s.CodeName, Ci);

        // Aggregate: one record per (procedure, sheet, target kind, address, defined name).
        var groups = sites
            .GroupBy(s => (Proc: $"{s.Module}.{s.Procedure}", s.SheetName, s.TargetKind, s.Address, s.DefinedNameRef), new KeyComparer())
            .Select(g =>
            {
                var modes = g.Select(s => s.Mode).Distinct(Ci).ToList();
                var mode = modes.Count > 1 ? "both" : modes[0].ToLowerInvariant();
                var first = g.First();
                return new ExcelVbaSheetAccess(
                    g.Key.Proc,
                    first.SheetName is null ? null : new ExcelVbaSheetRef(first.SheetName, codeNameBySheet.GetValueOrDefault(first.SheetName)),
                    new ExcelVbaAccessTarget(first.TargetKind, first.Address, first.DefinedNameRef),
                    mode,
                    g.Count(),
                    first.SheetName is null ? first.UnresolvedReason ?? "unknownSheet" : null);
            })
            .ToList();

        var rollup = sites.Where(s => s.SheetName is not null)
            .GroupBy(s => s.SheetName!, Ci)
            .Select(g => new ExcelVbaSheetUsage(
                g.Key,
                codeNameBySheet.GetValueOrDefault(g.Key),
                g.Where(s => Ci.Equals(s.Mode, "read")).Select(Fqn).Distinct(Ci).OrderBy(x => x, Ci).ToList(),
                g.Where(s => Ci.Equals(s.Mode, "write")).Select(Fqn).Distinct(Ci).OrderBy(x => x, Ci).ToList(),
                g.Count(s => Ci.Equals(s.Mode, "read")),
                g.Count(s => Ci.Equals(s.Mode, "write"))))
            .OrderBy(u => u.Name, Ci)
            .ToList();

        var summary = new ExcelVbaSheetAccessSummary(
            sites.Count,
            sites.Count(s => s.SheetName is not null),
            sites.Count(s => s.SheetName is null),
            rollup.Count,
            sites.Select(Fqn).Distinct(Ci).Count());

        IEnumerable<ExcelVbaSheetAccess> access = groups;
        IEnumerable<ExcelVbaSheetUsage> usage = rollup;
        if (moduleFilter is not null)
            access = access.Where(a => Ci.Equals(a.Procedure.Split('.', 2)[0], moduleFilter));
        if (sheetFilter is not null)
        {
            access = access.Where(a => a.Sheet is not null && Ci.Equals(a.Sheet.Name, sheetFilter));
            usage = usage.Where(u => Ci.Equals(u.Name, sheetFilter));
        }
        if (!includeUnresolved)
            access = access.Where(a => a.Sheet is not null);

        var accessList = access
            .OrderBy(a => a.Procedure, Ci)
            .ThenBy(a => a.Sheet?.Name ?? "~", Ci)
            .ThenBy(a => a.Target.Address ?? a.Target.DefinedName ?? "", Ci)
            .ToList();

        // ponytail: includeRecords=false is the cheap first call on a big workbook (Air: 672 records = 114 KB, rollup 9 KB).
        return new ExcelVbaSheetAccessResult(path, true, summary,
            includeRecords ? accessList.Take(maxRecords).ToList() : [], usage.ToList(),
            includeRecords && accessList.Count > maxRecords);
    }

    private static string Fqn(VbaSheetAccessResolver.AccessSite s) => $"{s.Module}.{s.Procedure}";

    private sealed class KeyComparer : IEqualityComparer<(string Proc, string? SheetName, string TargetKind, string? Address, string? DefinedNameRef)>
    {
        public bool Equals((string Proc, string? SheetName, string TargetKind, string? Address, string? DefinedNameRef) x,
                           (string Proc, string? SheetName, string TargetKind, string? Address, string? DefinedNameRef) y) =>
            Ci.Equals(x.Proc, y.Proc) && Ci.Equals(x.SheetName, y.SheetName) && Ci.Equals(x.TargetKind, y.TargetKind)
            && Ci.Equals(x.Address, y.Address) && Ci.Equals(x.DefinedNameRef, y.DefinedNameRef);

        public int GetHashCode((string Proc, string? SheetName, string TargetKind, string? Address, string? DefinedNameRef) k) =>
            HashCode.Combine(Ci.GetHashCode(k.Proc), k.SheetName is null ? 0 : Ci.GetHashCode(k.SheetName),
                Ci.GetHashCode(k.TargetKind), k.Address is null ? 0 : Ci.GetHashCode(k.Address),
                k.DefinedNameRef is null ? 0 : Ci.GetHashCode(k.DefinedNameRef));
    }
}
