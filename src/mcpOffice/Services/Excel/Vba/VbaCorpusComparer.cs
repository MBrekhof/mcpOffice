using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

/// <summary>
/// Groups procedures across workbooks: identical bodies by hash, then same-named near-duplicates by
/// line similarity. Pure: takes already-read projects so tests can feed in-memory ones.
/// </summary>
internal static class VbaCorpusComparer
{
    public sealed record WorkbookInput(string Path, ExcelVbaProject? Project, string? Error);

    private const double NearDuplicateThreshold = 0.9;
    private const double SharedModuleRatio = 0.5;
    // ponytail: bodies shorter than this are noise (empty handlers, one-line wrappers); raise if the corpus disagrees.
    private const int MinBodyLines = 3;
    private static readonly StringComparer Ci = StringComparer.OrdinalIgnoreCase;

    private sealed record Proc(string Workbook, string Module, string Name, string Hash, IReadOnlyList<string> Body);

    public static ExcelVbaCorpusResult Compare(IReadOnlyList<WorkbookInput> inputs, int minOccurrences, int maxProcedures, bool includeNearDuplicates)
    {
        minOccurrences = Math.Max(2, minOccurrences);
        var workbooks = new List<ExcelVbaCorpusWorkbook>();
        var procs = new List<Proc>();

        foreach (var input in inputs)
        {
            if (input.Project is null || !input.Project.HasVbaProject)
            {
                workbooks.Add(new ExcelVbaCorpusWorkbook(input.Path, input.Project?.HasVbaProject ?? false, 0, 0, input.Error));
                continue;
            }
            int procedureCount = 0;
            foreach (var m in input.Project.Modules)
            {
                if (string.IsNullOrEmpty(m.Code)) continue;
                var lines = VbaLineCleaner.Clean(m.Code);
                foreach (var sp in VbaProcedureScanner.Scan(m.Kind, m.Name, lines))
                {
                    procedureCount++;
                    var body = VbaProcedureHasher.Normalize(lines, sp.CleanedLineStartIndex, sp.CleanedLineEndIndex);
                    if (body.Count < MinBodyLines) continue;
                    procs.Add(new Proc(input.Path, m.Name, sp.Procedure.Name, VbaProcedureHasher.Hash(body), body));
                }
            }
            workbooks.Add(new ExcelVbaCorpusWorkbook(input.Path, true, input.Project.Modules.Count, procedureCount, null));
        }

        var shared = new List<ExcelVbaSharedProcedure>();
        var inGroup = new HashSet<Proc>();

        // Tier 1: identical bodies.
        foreach (var g in procs.GroupBy(p => p.Hash, StringComparer.Ordinal))
        {
            var members = g.ToList();
            if (members.Select(p => p.Workbook).Distinct(Ci).Count() < minOccurrences) continue;
            foreach (var p in members) inGroup.Add(p);
            shared.Add(new ExcelVbaSharedProcedure("identical", MostCommonName(members), members[0].Body.Count,
                members.Select(p => new ExcelVbaProcedureOccurrence(p.Workbook, p.Module, p.Name, 1.0)).ToList()));
        }

        // Tier 2: same name, different body, ≥ threshold similar to the longest member.
        if (includeNearDuplicates)
        {
            foreach (var g in procs.Where(p => !inGroup.Contains(p)).GroupBy(p => p.Name, Ci))
            {
                var candidates = g.ToList();
                if (candidates.Select(p => p.Workbook).Distinct(Ci).Count() < minOccurrences) continue;
                var anchor = candidates.OrderByDescending(p => p.Body.Count).First();
                var kept = candidates
                    .Select(p => (Proc: p, Sim: VbaProcedureHasher.Similarity(anchor.Body, p.Body)))
                    .Where(x => x.Sim >= NearDuplicateThreshold)
                    .ToList();
                if (kept.Select(x => x.Proc.Workbook).Distinct(Ci).Count() < minOccurrences) continue;
                foreach (var x in kept) inGroup.Add(x.Proc);
                shared.Add(new ExcelVbaSharedProcedure("nearDuplicate", anchor.Name, anchor.Body.Count,
                    kept.Select(x => new ExcelVbaProcedureOccurrence(x.Proc.Workbook, x.Proc.Module, x.Proc.Name, Math.Round(x.Sim, 3))).ToList()));
            }
        }

        // Shared modules: same module name in ≥ minOccurrences workbooks, mostly shared procedures.
        var sharedModules = new List<ExcelVbaSharedModule>();
        foreach (var g in procs.GroupBy(p => p.Module, Ci))
        {
            var perWorkbook = g.GroupBy(p => p.Workbook, Ci).ToList();
            if (perWorkbook.Count < minOccurrences) continue;
            var ratio = (double)g.Count(inGroup.Contains) / g.Count();
            if (ratio < SharedModuleRatio) continue;
            sharedModules.Add(new ExcelVbaSharedModule(g.Key, perWorkbook.Select(w => w.Key).OrderBy(w => w, Ci).ToList(), Math.Round(ratio, 3)));
        }

        var ordered = shared
            .OrderByDescending(s => s.Occurrences.Count)
            .ThenBy(s => s.Name, Ci)
            .ToList();

        var summary = new ExcelVbaCorpusSummary(
            workbooks.Count,
            workbooks.Sum(w => w.ProcedureCount),
            inGroup.Count,
            shared.Count(s => s.Tier == "identical"),
            shared.Count(s => s.Tier == "nearDuplicate"),
            sharedModules.Count);

        return new ExcelVbaCorpusResult(
            workbooks,
            summary,
            ordered.Take(maxProcedures).ToList(),
            sharedModules.OrderBy(m => m.Module, Ci).ToList(),
            ordered.Count > maxProcedures);
    }

    private static string MostCommonName(List<Proc> members) =>
        members.GroupBy(p => p.Name, Ci).OrderByDescending(g => g.Count()).ThenBy(g => g.Key, Ci).First().Key;
}
