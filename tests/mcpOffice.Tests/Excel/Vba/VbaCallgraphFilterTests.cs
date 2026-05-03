using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaCallgraphFilterTests
{
    // Helper: build a minimal ExcelVbaAnalysis with the given modules + edges.
    private static ExcelVbaAnalysis Analysis(
        IEnumerable<(string Module, string Kind, string Name, bool IsHandler)> procs,
        IEnumerable<(string From, string To, bool Resolved)> edges)
    {
        var byModule = procs
            .GroupBy(p => (p.Module, p.Kind))
            .Select(g => new ExcelVbaModuleAnalysis(
                g.Key.Module,
                g.Key.Kind,
                Parsed: true,
                Reason: null,
                Procedures: g.Select(p => new ExcelVbaProcedure(
                    Name: p.Name,
                    FullyQualifiedName: $"{p.Module}.{p.Name}",
                    Kind: "Sub",
                    Scope: null,
                    Parameters: Array.Empty<ExcelVbaParameter>(),
                    ReturnType: null,
                    LineStart: 1,
                    LineEnd: 2,
                    IsEventHandler: p.IsHandler,
                    EventTarget: null)).ToList()))
            .ToList();

        var callEdges = edges.Select(e => new ExcelVbaCallEdge(
            From: e.From,
            To: e.To,
            Resolved: e.Resolved,
            Site: new ExcelVbaSiteRef(
                Module: e.From.Split('.')[0],
                Procedure: e.From.Split('.')[1],
                Line: 1))).ToList();

        var procedureCount = byModule.Sum(m => m.Procedures.Count);
        var handlerCount = byModule.Sum(m => m.Procedures.Count(p => p.IsEventHandler));

        return new ExcelVbaAnalysis(
            HasVbaProject: true,
            Summary: new ExcelVbaAnalysisSummary(
                ModuleCount: byModule.Count,
                ParsedModuleCount: byModule.Count,
                UnparsedModuleCount: 0,
                ProcedureCount: procedureCount,
                EventHandlerCount: handlerCount,
                CallEdgeCount: callEdges.Count,
                ObjectModelReferenceCount: 0,
                DependencyCount: 0),
            Modules: byModule,
            CallGraph: callEdges,
            References: null);
    }

    [Fact]
    public void No_filter_returns_every_procedure_as_a_node()
    {
        var a = Analysis(
            procs: new[]
            {
                ("ModA", "standardModule", "P1", false),
                ("ModA", "standardModule", "P2", false),
                ("ModB", "standardModule", "Q1", false),
            },
            edges: new[]
            {
                ("ModA.P1", "ModA.P2", true),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

        Assert.Equal(3, result.Nodes.Count);
        Assert.Contains(result.Nodes, n => n.Id == "ModA.P1");
        Assert.Contains(result.Nodes, n => n.Id == "ModA.P2");
        Assert.Contains(result.Nodes, n => n.Id == "ModB.Q1");
        Assert.Single(result.Edges);
        Assert.Equal("ModA.P1", result.Edges[0].FromId);
        Assert.Equal("ModA.P2", result.Edges[0].ToId);
        Assert.True(result.Edges[0].Resolved);
    }

    [Fact]
    public void No_vba_project_returns_empty()
    {
        var empty = new ExcelVbaAnalysis(
            HasVbaProject: false,
            Summary: new ExcelVbaAnalysisSummary(0, 0, 0, 0, 0, 0, 0, 0),
            Modules: null,
            CallGraph: null,
            References: null);

        var result = VbaCallgraphFilter.Apply(empty, new CallgraphFilterOptions());

        Assert.Empty(result.Nodes);
        Assert.Empty(result.Edges);
    }
}
