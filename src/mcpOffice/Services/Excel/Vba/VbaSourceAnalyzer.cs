using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class VbaSourceAnalyzer
{
    private const int MaxLinesPerModule = 5000;

    public static ExcelVbaAnalysis Analyze(
        ExcelVbaProject project,
        bool includeProcedures,
        bool includeCallGraph,
        bool includeReferences)
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
        foreach (var (moduleName, _, lines, procs, parsed, _) in perModule)
        {
            if (!parsed) continue;
            VbaReferenceCollector.Collect(moduleName, lines, procs, omRefs, deps);
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

        return new ExcelVbaAnalysis(
            HasVbaProject: true,
            Summary: summary,
            Modules: includeProcedures ? moduleAnalyses : null,
            CallGraph: includeCallGraph ? edges : null,
            References: includeReferences ? new ExcelVbaReferences(omRefs, deps) : null);
    }
}
