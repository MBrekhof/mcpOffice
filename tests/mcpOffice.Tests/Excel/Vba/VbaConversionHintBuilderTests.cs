using McpOffice.Models;
using McpOffice.Services.Excel.Vba;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class VbaConversionHintBuilderTests
{
    private static ExcelVbaProcedure Proc(
        string module, string name,
        bool isEventHandler = false, string? scope = "Public") =>
        new(name, $"{module}.{name}", "Sub", scope,
            Array.Empty<ExcelVbaParameter>(), null, 1, 2, isEventHandler, null);

    private static ExcelVbaModuleAnalysis Mod(string name, string kind, params ExcelVbaProcedure[] procs) =>
        new(name, kind, true, null, procs);

    private static ExcelVbaCallEdge Edge(string from, string to, bool resolved = true) =>
        new(from, to, resolved,
            new ExcelVbaSiteRef(from.Split('.')[0], from.Split('.')[1], 1));

    private static ExcelVbaAnalysis MakeAnalysis(
        IReadOnlyList<ExcelVbaModuleAnalysis> modules,
        IReadOnlyList<ExcelVbaCallEdge> callGraph,
        IReadOnlyList<ExcelVbaObjectModelRef>? om = null,
        IReadOnlyList<ExcelVbaDependency>? deps = null)
    {
        var refs = new ExcelVbaReferences(om ?? Array.Empty<ExcelVbaObjectModelRef>(),
                                          deps ?? Array.Empty<ExcelVbaDependency>());
        var procCount = modules.Sum(m => m.Procedures.Count);
        var handlerCount = modules.Sum(m => m.Procedures.Count(p => p.IsEventHandler));
        var summary = new ExcelVbaAnalysisSummary(
            modules.Count, modules.Count, 0, procCount, handlerCount,
            callGraph.Count, refs.ObjectModel.Count, refs.Dependencies.Count);
        return new ExcelVbaAnalysis(true, summary, modules, callGraph, refs);
    }

    [Fact]
    public void Builder_emits_one_hint_per_procedure()
    {
        var analysis = MakeAnalysis(
            new[]
            {
                Mod("Module1", "standard", Proc("Module1", "A"), Proc("Module1", "B")),
                Mod("Module2", "standard", Proc("Module2", "C"))
            },
            Array.Empty<ExcelVbaCallEdge>());

        var result = VbaConversionHintBuilder.Build(analysis, moduleName: null, targetParadigm: null);
        Assert.Equal(3, result.ProcedureHints.Count);
        Assert.Equal(3, result.Summary.TotalProcedures);
        Assert.Equal(3, result.Summary.HintedProcedures);
        Assert.Equal(2, result.Summary.ModuleCount);
    }

    [Fact]
    public void Builder_filters_hints_by_moduleName_case_insensitive()
    {
        var analysis = MakeAnalysis(
            new[]
            {
                Mod("Module1", "standard", Proc("Module1", "A")),
                Mod("Module2", "standard", Proc("Module2", "C"))
            },
            Array.Empty<ExcelVbaCallEdge>());

        var result = VbaConversionHintBuilder.Build(analysis, moduleName: "MODULE1", targetParadigm: null);
        Assert.Single(result.ProcedureHints);
        Assert.Equal("Module1", result.ProcedureHints[0].Module);
        Assert.Equal(1, result.Summary.HintedProcedures);
        Assert.Equal(2, result.Summary.TotalProcedures);
    }

    [Fact]
    public void Builder_throws_module_not_found_for_unknown_filter()
    {
        var analysis = MakeAnalysis(
            new[] { Mod("Module1", "standard", Proc("Module1", "A")) },
            Array.Empty<ExcelVbaCallEdge>());
        var act = () => VbaConversionHintBuilder.Build(analysis, moduleName: "Nope", targetParadigm: null);
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("[module_not_found]", ex.Message);
    }

    [Fact]
    public void Builder_throws_unsupported_paradigm()
    {
        var analysis = MakeAnalysis(
            new[] { Mod("Module1", "standard", Proc("Module1", "A")) },
            Array.Empty<ExcelVbaCallEdge>());
        var act = () => VbaConversionHintBuilder.Build(analysis, moduleName: null, targetParadigm: "blazor");
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("[unsupported_paradigm]", ex.Message);
        Assert.Contains("blazor", ex.Message);
    }

    [Fact]
    public void Builder_emits_csharpSuggestion_only_when_paradigm_provided()
    {
        var analysis = MakeAnalysis(
            new[] { Mod("Module1", "standard", Proc("Module1", "A")) },
            Array.Empty<ExcelVbaCallEdge>());

        var without = VbaConversionHintBuilder.Build(analysis, null, null);
        Assert.Null(without.ProcedureHints[0].CsharpSuggestion);

        var with = VbaConversionHintBuilder.Build(analysis, null, "classLibrary");
        Assert.NotNull(with.ProcedureHints[0].CsharpSuggestion);
    }

    [Fact]
    public void Builder_coupling_always_whole_workbook_even_with_module_filter()
    {
        var analysis = MakeAnalysis(
            new[]
            {
                Mod("A", "standard", Proc("A", "f")),
                Mod("B", "standard", Proc("B", "g")),
                Mod("C", "standard", Proc("C", "h"))
            },
            new[] { Edge("A.f", "B.g"), Edge("B.g", "C.h") });

        var result = VbaConversionHintBuilder.Build(analysis, moduleName: "A", targetParadigm: null);
        Assert.Equal(3, result.ModuleCoupling.Count);  // all modules present
        Assert.Equal(2, result.CouplingPairs.Count);   // A→B, B→C
    }

    [Fact]
    public void Builder_caller_callee_counts_match_call_graph()
    {
        var analysis = MakeAnalysis(
            new[]
            {
                Mod("M", "standard",
                    Proc("M", "Caller", scope: "Public"),
                    Proc("M", "Callee", scope: "Public"))
            },
            new[] { Edge("M.Caller", "M.Callee") });

        var result = VbaConversionHintBuilder.Build(analysis, null, null);
        var caller = result.ProcedureHints.Single(p => p.ProcedureName == "Caller");
        var callee = result.ProcedureHints.Single(p => p.ProcedureName == "Callee");
        Assert.Equal(0, caller.CallerCount);
        Assert.Equal(1, caller.CalleeCount);
        Assert.Equal(1, callee.CallerCount);
        Assert.Equal(0, callee.CalleeCount);
    }

    [Fact]
    public void Builder_returns_empty_hints_when_no_vba_project()
    {
        var empty = new ExcelVbaAnalysis(
            HasVbaProject: false,
            Summary: new ExcelVbaAnalysisSummary(0, 0, 0, 0, 0, 0, 0, 0),
            Modules: null, CallGraph: null, References: null);
        var result = VbaConversionHintBuilder.Build(empty, null, null);
        Assert.Empty(result.ProcedureHints);
        Assert.Empty(result.ModuleCoupling);
        Assert.Empty(result.CouplingPairs);
    }
}
