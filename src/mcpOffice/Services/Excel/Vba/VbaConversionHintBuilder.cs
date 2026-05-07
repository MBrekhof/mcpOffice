using System.Diagnostics;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class VbaConversionHintBuilder
{
    public static ConversionHints Build(
        ExcelVbaAnalysis analysis,
        string? moduleName,
        string? targetParadigm)
    {
        var sw = Stopwatch.StartNew();

        if (targetParadigm is not null &&
            !ParadigmOverlayApplier.SupportedParadigms.Contains(targetParadigm))
        {
            throw ToolError.UnsupportedParadigm(targetParadigm, ParadigmOverlayApplier.SupportedParadigms);
        }

        if (!analysis.HasVbaProject || analysis.Modules is null)
        {
            return new ConversionHints(
                Summary: new ConversionHintsSummary(0, 0, 0, targetParadigm, sw.ElapsedMilliseconds),
                ProcedureHints: Array.Empty<ProcedureHint>(),
                ModuleCoupling: Array.Empty<ModuleCoupling>(),
                CouplingPairs: Array.Empty<CouplingPair>());
        }

        var modules = analysis.Modules!;
        var callGraph = analysis.CallGraph ?? Array.Empty<ExcelVbaCallEdge>();
        var objectModel = analysis.References?.ObjectModel ?? Array.Empty<ExcelVbaObjectModelRef>();
        var dependencies = analysis.References?.Dependencies ?? Array.Empty<ExcelVbaDependency>();

        // Resolve moduleName (case-insensitive) to canonical name; throw if unknown.
        string? canonicalFilter = null;
        if (!string.IsNullOrWhiteSpace(moduleName))
        {
            var match = modules.FirstOrDefault(m =>
                string.Equals(m.Name, moduleName, StringComparison.OrdinalIgnoreCase));
            if (match is null)
            {
                throw ToolError.ModuleNotFound(moduleName, modules.Select(m => m.Name));
            }
            canonicalFilter = match.Name;
        }

        // Coupling — always whole-workbook.
        var moduleNamesAndKinds = modules.Select(m => (m.Name, m.Kind)).ToList();
        var couplingResult = CouplingComputer.Compute(moduleNamesAndKinds, callGraph);

        // Hints — filtered by moduleName when set.
        var hints = new List<ProcedureHint>();
        int totalProcedures = 0;

        foreach (var mod in modules)
        {
            totalProcedures += mod.Procedures.Count;

            if (canonicalFilter is not null && !string.Equals(mod.Name, canonicalFilter, StringComparison.Ordinal))
                continue;

            foreach (var proc in mod.Procedures)
            {
                var axes = AxisClassifier.Classify(proc, mod.Kind, callGraph, objectModel, dependencies);
                int callerCount = callGraph.Count(e =>
                    string.Equals(e.To, proc.FullyQualifiedName, StringComparison.OrdinalIgnoreCase));
                int calleeCount = callGraph.Count(e =>
                    string.Equals(e.From, proc.FullyQualifiedName, StringComparison.OrdinalIgnoreCase));

                CSharpSuggestion? suggestion = null;
                if (targetParadigm is not null)
                {
                    suggestion = ParadigmOverlayApplier.Apply(
                        mod.Name, proc.Name, proc.Scope, axes, targetParadigm);
                }

                var rationale = BuildRationale(proc, axes, suggestion);

                hints.Add(new ProcedureHint(
                    Module: mod.Name,
                    ProcedureName: proc.Name,
                    Kind: proc.Kind,
                    IsEventHandler: proc.IsEventHandler,
                    ParamCount: proc.Parameters.Count,
                    CallerCount: callerCount,
                    CalleeCount: calleeCount,
                    Axes: axes,
                    Rationale: rationale,
                    CsharpSuggestion: suggestion));
            }
        }

        var summary = new ConversionHintsSummary(
            TotalProcedures: totalProcedures,
            HintedProcedures: hints.Count,
            ModuleCount: modules.Count,
            TargetParadigm: targetParadigm,
            WallTimeMs: sw.ElapsedMilliseconds);

        return new ConversionHints(summary, hints, couplingResult.Coupling, couplingResult.Pairs);
    }

    private static string BuildRationale(
        ExcelVbaProcedure proc, ProcedureAxes axes, CSharpSuggestion? suggestion)
    {
        var parts = new List<string>
        {
            $"{axes.Purity} {(axes.Shape ?? "")}".Trim(),
            $"trigger={axes.Trigger}",
            $"params={proc.Parameters.Count}"
        };
        if (axes.Dependencies.Count > 0)
            parts.Add($"deps=[{string.Join(",", axes.Dependencies)}]");

        var baseLine = string.Join("; ", parts);

        if (suggestion is null) return baseLine;

        var paradigmHint = suggestion.TargetType == "requiresManualReview"
            ? $"Manual review required: {string.Join(", ", suggestion.Blockers)}."
            : $"Suggested as {suggestion.TargetType} on {suggestion.SuggestedClassName}.{suggestion.SuggestedMethodName}.";

        return $"{baseLine} — {paradigmHint}";
    }
}
