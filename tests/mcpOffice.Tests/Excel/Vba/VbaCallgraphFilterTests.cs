using McpOffice.Models;
using McpOffice.Services.Excel.Vba;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class VbaCallgraphFilterTests
{
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

    // ----- Task 4: no-filter pass-through -----

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

    // ----- Task 5: moduleName direct-neighbour expansion + module_not_found -----

    [Fact]
    public void Module_filter_includes_module_procedures_and_direct_neighbours()
    {
        var a = Analysis(
            procs: new[]
            {
                ("ModA", "standardModule", "P1", false),
                ("ModB", "standardModule", "Q1", false),
                ("ModC", "standardModule", "R1", false),
            },
            edges: new[]
            {
                ("ModA.P1", "ModB.Q1", true),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "ModA"));

        Assert.Equal(2, result.Nodes.Count);
        Assert.Contains(result.Nodes, n => n.Id == "ModA.P1");
        Assert.Contains(result.Nodes, n => n.Id == "ModB.Q1");
        Assert.DoesNotContain(result.Nodes, n => n.Id == "ModC.R1");
        Assert.Single(result.Edges);
    }

    [Fact]
    public void Module_filter_pulls_in_callers_too()
    {
        var a = Analysis(
            procs: new[]
            {
                ("ModA", "standardModule", "P1", false),
                ("ModB", "standardModule", "Q1", false),
            },
            edges: new[]
            {
                ("ModB.Q1", "ModA.P1", true),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "ModA"));

        Assert.Equal(2, result.Nodes.Count);
        Assert.Single(result.Edges);
    }

    [Fact]
    public void Module_filter_is_case_insensitive()
    {
        var a = Analysis(
            procs: new[] { ("ModA", "standardModule", "P1", false) },
            edges: Array.Empty<(string, string, bool)>());

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "moda"));

        Assert.Single(result.Nodes);
    }

    [Fact]
    public void Module_filter_unknown_throws_module_not_found()
    {
        var a = Analysis(
            procs: new[] { ("ModA", "standardModule", "P1", false) },
            edges: Array.Empty<(string, string, bool)>());

        var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "Nope"));
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("module_not_found", ex.Message);
        Assert.Contains("ModA", ex.Message);
    }

    // ----- Task 6: focal-procedure BFS + procedure_not_found + invalid direction -----

    [Fact]
    public void Focal_procedure_callees_only_depth_1()
    {
        var a = Analysis(
            procs: new[]
            {
                ("M", "standardModule", "P1", false),
                ("M", "standardModule", "P2", false),
                ("M", "standardModule", "P3", false),
                ("M", "standardModule", "P4", false),
            },
            edges: new[]
            {
                ("M.P1", "M.P2", true),
                ("M.P2", "M.P3", true),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ModuleName: "M",
            ProcedureName: "P1",
            Depth: 1,
            Direction: "callees"));

        Assert.Equal(2, result.Nodes.Count);
        Assert.Contains(result.Nodes, n => n.Id == "M.P1");
        Assert.Contains(result.Nodes, n => n.Id == "M.P2");
        Assert.DoesNotContain(result.Nodes, n => n.Id == "M.P3");
        Assert.DoesNotContain(result.Nodes, n => n.Id == "M.P4");
    }

    [Fact]
    public void Focal_procedure_callees_depth_2_pulls_in_grandchildren()
    {
        var a = Analysis(
            procs: new[]
            {
                ("M", "standardModule", "P1", false),
                ("M", "standardModule", "P2", false),
                ("M", "standardModule", "P3", false),
            },
            edges: new[]
            {
                ("M.P1", "M.P2", true),
                ("M.P2", "M.P3", true),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ModuleName: "M",
            ProcedureName: "P1",
            Depth: 2,
            Direction: "callees"));

        Assert.Equal(3, result.Nodes.Count);
    }

    [Fact]
    public void Focal_procedure_callers_walks_inbound_edges()
    {
        var a = Analysis(
            procs: new[]
            {
                ("M", "standardModule", "P1", false),
                ("M", "standardModule", "P2", false),
                ("M", "standardModule", "P3", false),
                ("M", "standardModule", "P4", false),
            },
            edges: new[]
            {
                ("M.P2", "M.P1", true),
                ("M.P3", "M.P2", true),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ModuleName: "M",
            ProcedureName: "P1",
            Depth: 2,
            Direction: "callers"));

        Assert.Equal(3, result.Nodes.Count);
        Assert.DoesNotContain(result.Nodes, n => n.Id == "M.P4");
    }

    [Fact]
    public void Focal_procedure_both_unions_callees_and_callers()
    {
        var a = Analysis(
            procs: new[]
            {
                ("M", "standardModule", "P0", false),
                ("M", "standardModule", "P1", false),
                ("M", "standardModule", "P2", false),
                ("M", "standardModule", "P3", false),
            },
            edges: new[]
            {
                ("M.P0", "M.P1", true),
                ("M.P1", "M.P2", true),
                ("M.P3", "M.P1", true),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ModuleName: "M",
            ProcedureName: "P1",
            Depth: 1,
            Direction: "both"));

        Assert.Equal(4, result.Nodes.Count);
    }

    [Fact]
    public void Focal_procedure_depth_zero_returns_just_the_seed()
    {
        var a = Analysis(
            procs: new[]
            {
                ("M", "standardModule", "P1", false),
                ("M", "standardModule", "P2", false),
            },
            edges: new[] { ("M.P1", "M.P2", true) });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ModuleName: "M",
            ProcedureName: "P1",
            Depth: 0,
            Direction: "both"));

        Assert.Single(result.Nodes);
        Assert.Equal("M.P1", result.Nodes[0].Id);
        Assert.Empty(result.Edges);
    }

    [Fact]
    public void ProcedureName_unknown_throws_procedure_not_found()
    {
        var a = Analysis(
            procs: new[]
            {
                ("M", "standardModule", "P1", false),
                ("M", "standardModule", "P2", false),
            },
            edges: Array.Empty<(string, string, bool)>());

        var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ModuleName: "M",
            ProcedureName: "Nope"));
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("procedure_not_found", ex.Message);
        Assert.Contains("Nope", ex.Message);
        Assert.Contains("P1", ex.Message);
        Assert.Contains("P2", ex.Message);
    }

    [Fact]
    public void ProcedureName_is_case_insensitive_within_module()
    {
        var a = Analysis(
            procs: new[] { ("M", "standardModule", "ReadExports", false) },
            edges: Array.Empty<(string, string, bool)>());

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ModuleName: "M",
            ProcedureName: "readexports",
            Depth: 0));

        Assert.Single(result.Nodes);
    }

    [Fact]
    public void Direction_unknown_value_throws_invalid_render_option()
    {
        var a = Analysis(
            procs: new[] { ("M", "standardModule", "P1", false) },
            edges: Array.Empty<(string, string, bool)>());

        var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ModuleName: "M",
            ProcedureName: "P1",
            Direction: "sideways"));
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("invalid_render_option", ex.Message);
        Assert.Contains("direction", ex.Message);
        Assert.Contains("sideways", ex.Message);
    }

    // ----- Task 7: procedureName without moduleName -----

    [Fact]
    public void ProcedureName_without_moduleName_throws_invalid_render_option()
    {
        var a = Analysis(
            procs: new[] { ("M", "standardModule", "P1", false) },
            edges: Array.Empty<(string, string, bool)>());

        var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(
            ProcedureName: "P1"));
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("invalid_render_option", ex.Message);
        Assert.Contains("procedureName", ex.Message);
        Assert.Contains("requires moduleName", ex.Message);
    }

    // ----- Task 8: external (unresolved) callees -----

    [Fact]
    public void Unresolved_callees_become_single_external_node_per_name()
    {
        var a = Analysis(
            procs: new[]
            {
                ("M", "standardModule", "P1", false),
                ("M", "standardModule", "P2", false),
            },
            edges: new[]
            {
                ("M.P1", "MsgBox", false),
                ("M.P2", "MsgBox", false),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

        var externals = result.Nodes.Where(n => n.IsExternal).ToList();
        Assert.Single(externals);
        Assert.Equal("MsgBox", externals[0].Label);
        Assert.True(externals[0].IsExternal);
        Assert.Null(externals[0].Module);

        var externalEdges = result.Edges.Where(e => e.ToId == externals[0].Id).ToList();
        Assert.Equal(2, externalEdges.Count);
        Assert.All(externalEdges, e => Assert.False(e.Resolved));
    }

    [Fact]
    public void Distinct_unresolved_names_get_distinct_external_nodes()
    {
        var a = Analysis(
            procs: new[] { ("M", "standardModule", "P1", false) },
            edges: new[]
            {
                ("M.P1", "MsgBox", false),
                ("M.P1", "CreateObject", false),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

        var externals = result.Nodes.Where(n => n.IsExternal).ToList();
        Assert.Equal(2, externals.Count);
        Assert.Contains(externals, e => e.Label == "MsgBox");
        Assert.Contains(externals, e => e.Label == "CreateObject");
    }

    [Fact]
    public void External_nodes_appear_only_when_caller_is_in_filtered_set()
    {
        var a = Analysis(
            procs: new[]
            {
                ("ModA", "standardModule", "P1", false),
                ("ModB", "standardModule", "Q1", false),
            },
            edges: new[]
            {
                ("ModA.P1", "MsgBox", false),
                ("ModB.Q1", "MsgBox", false),
            });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "ModA"));

        var externals = result.Nodes.Where(n => n.IsExternal).ToList();
        Assert.Single(externals);
        var externalEdges = result.Edges.Where(e => e.ToId == externals[0].Id).ToList();
        Assert.Single(externalEdges);
        Assert.Equal("ModA.P1", externalEdges[0].FromId);
    }

    // ----- Task 9: orphan classification -----

    [Fact]
    public void Orphan_procedure_with_no_edges_marked_isOrphan()
    {
        var a = Analysis(
            procs: new[]
            {
                ("M", "standardModule", "Connected1", false),
                ("M", "standardModule", "Connected2", false),
                ("M", "standardModule", "Lonely", false),
            },
            edges: new[] { ("M.Connected1", "M.Connected2", true) });

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

        var lonely = Assert.Single(result.Nodes, n => n.Id == "M.Lonely");
        Assert.True(lonely.IsOrphan);
        Assert.False(result.Nodes.Single(n => n.Id == "M.Connected1").IsOrphan);
    }

    [Fact]
    public void Event_handler_with_no_edges_is_NOT_marked_orphan()
    {
        var a = Analysis(
            procs: new[]
            {
                ("ThisWorkbook", "documentModule", "Workbook_Open", true),
            },
            edges: Array.Empty<(string, string, bool)>());

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

        var node = Assert.Single(result.Nodes);
        Assert.True(node.IsEventHandler);
        Assert.False(node.IsOrphan);
    }

    [Fact]
    public void Orphan_classification_is_per_filtered_view_not_per_workbook()
    {
        var a = Analysis(
            procs: new[]
            {
                ("ModA", "standardModule", "P1", false),
                ("ModA", "standardModule", "X", false),
                ("ModB", "standardModule", "Q1", false),
            },
            edges: new[] { ("ModA.P1", "ModB.Q1", true) });

        var moduleResult = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(ModuleName: "ModA"));
        var noFilterResult = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions());

        Assert.True(moduleResult.Nodes.Single(n => n.Id == "ModA.X").IsOrphan);
        Assert.True(noFilterResult.Nodes.Single(n => n.Id == "ModA.X").IsOrphan);
        Assert.False(noFilterResult.Nodes.Single(n => n.Id == "ModA.P1").IsOrphan);
    }

    // ----- Task 10: maxNodes cap -----

    [Fact]
    public void Exceeds_maxNodes_throws_graph_too_large()
    {
        var procs = Enumerable.Range(0, 5)
            .Select(i => ("M", "standardModule", $"P{i}", false))
            .ToArray();
        var a = Analysis(procs, Array.Empty<(string, string, bool)>());

        var act = () => VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(MaxNodes: 3));
        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("graph_too_large", ex.Message);
        Assert.Contains("5", ex.Message);
        Assert.Contains("3", ex.Message);
    }

    [Fact]
    public void Equal_to_maxNodes_does_not_throw()
    {
        var procs = Enumerable.Range(0, 3)
            .Select(i => ("M", "standardModule", $"P{i}", false))
            .ToArray();
        var a = Analysis(procs, Array.Empty<(string, string, bool)>());

        var result = VbaCallgraphFilter.Apply(a, new CallgraphFilterOptions(MaxNodes: 3));
        Assert.Equal(3, result.Nodes.Count);
    }
}
