using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class VbaSourceAnalyzer
{
    private const int MaxLinesPerModule = 5000;

    public static ExcelVbaAnalysis Analyze(
        ExcelVbaProject project,
        bool includeProcedures,
        bool includeCallGraph,
        bool includeReferences,
        string? moduleName = null)
    {
        if (!project.HasVbaProject)
        {
            return new ExcelVbaAnalysis(
                HasVbaProject: false,
                Summary: new ExcelVbaAnalysisSummary(0, 0, 0, 0, 0, 0, 0, 0),
                Modules: null,
                CallGraph: null,
                References: null);
        }

        var filter = string.IsNullOrWhiteSpace(moduleName) ? null : moduleName;
        if (filter is not null)
        {
            var match = project.Modules.FirstOrDefault(m =>
                string.Equals(m.Name, filter, StringComparison.OrdinalIgnoreCase));
            if (match is null)
            {
                throw ToolError.ModuleNotFound(filter, project.Modules.Select(m => m.Name));
            }
            filter = match.Name; // canonical casing for downstream comparisons
        }

        var perModule = new List<(string ModuleName, string ModuleKind, IReadOnlyList<CleanedLine> Lines, IReadOnlyList<ScannedProcedure> Procs, bool Parsed, string? Reason)>();
        var moduleAnalyses = new List<ExcelVbaModuleAnalysis>(project.Modules.Count);

        foreach (var m in project.Modules)
        {
            if (string.IsNullOrEmpty(m.Code))
            {
                moduleAnalyses.Add(new ExcelVbaModuleAnalysis(m.Name, m.Kind, false, "empty_source", []));
                perModule.Add((m.Name, m.Kind, [], [], false, "empty_source"));
                continue;
            }

            var cleaned = VbaLineCleaner.Clean(m.Code);
            if (cleaned.Count > MaxLinesPerModule)
            {
                moduleAnalyses.Add(new ExcelVbaModuleAnalysis(m.Name, m.Kind, false, "module_too_large", []));
                perModule.Add((m.Name, m.Kind, cleaned, [], false, "module_too_large"));
                continue;
            }

            var procs = VbaProcedureScanner.Scan(m.Kind, m.Name, cleaned);
            moduleAnalyses.Add(new ExcelVbaModuleAnalysis(
                m.Name, m.Kind, true, null, procs.Select(sp => sp.Procedure).ToList()));
            perModule.Add((m.Name, m.Kind, cleaned, procs, true, null));
        }

        // Call graph + references always built (cheap relative to extraction); we just decide whether to expose.
        var callModules = perModule.Where(p => p.Parsed)
            .Select(p => (p.ModuleName, p.Lines, p.Procs)).ToList();
        var edges = VbaCallGraphBuilder.Build(callModules);

        var omRefs = new List<ExcelVbaObjectModelRef>();
        var deps = new List<ExcelVbaDependency>();
        foreach (var (modName, _, lines, procs, parsed, _) in perModule)
        {
            if (!parsed) continue;
            VbaReferenceCollector.Collect(modName, lines, procs, omRefs, deps);
        }

        var procedureCount = moduleAnalyses.Sum(m => m.Procedures.Count);
        var eventHandlerCount = moduleAnalyses.Sum(m => m.Procedures.Count(p => p.IsEventHandler));
        var summary = new ExcelVbaAnalysisSummary(
            ModuleCount: project.Modules.Count,
            ParsedModuleCount: moduleAnalyses.Count(m => m.Parsed),
            UnparsedModuleCount: moduleAnalyses.Count(m => !m.Parsed),
            ProcedureCount: procedureCount,
            EventHandlerCount: eventHandlerCount,
            CallEdgeCount: edges.Count,
            ObjectModelReferenceCount: omRefs.Count,
            DependencyCount: deps.Count);

        IReadOnlyList<ExcelVbaModuleAnalysis>? modulesOut = null;
        IReadOnlyList<ExcelVbaCallEdge>? callGraphOut = null;
        ExcelVbaReferences? referencesOut = null;

        if (includeProcedures)
        {
            modulesOut = filter is null
                ? moduleAnalyses
                : moduleAnalyses.Where(m => m.Name == filter).ToList();
        }
        if (includeCallGraph)
        {
            callGraphOut = filter is null
                ? edges
                : edges.Where(e => InvolvesModule(e, filter)).ToList();
        }
        if (includeReferences)
        {
            referencesOut = filter is null
                ? new ExcelVbaReferences(omRefs, deps)
                : new ExcelVbaReferences(
                    omRefs.Where(r => r.Module == filter).ToList(),
                    deps.Where(d => d.Module == filter).ToList());
        }

        return new ExcelVbaAnalysis(
            HasVbaProject: true,
            Summary: summary,
            Modules: modulesOut,
            CallGraph: callGraphOut,
            References: referencesOut);
    }

    private static bool InvolvesModule(ExcelVbaCallEdge edge, string moduleName)
    {
        if (edge.Site.Module == moduleName) return true;
        if (edge.Resolved && edge.To.StartsWith(moduleName + ".", StringComparison.Ordinal)) return true;
        return false;
    }
}
