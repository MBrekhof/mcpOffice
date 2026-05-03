using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaSourceAnalyzerTests
{
    private static ExcelVbaProject Project(params (string Name, string Kind, string Code)[] modules) =>
        new(true, modules.Select(m => new ExcelVbaModule(m.Name, m.Kind, m.Code.Split('\n').Length, m.Code)).ToList());

    [Fact]
    public void HasVbaProject_false_returns_zeroed_summary()
    {
        var result = VbaSourceAnalyzer.Analyze(
            new ExcelVbaProject(false, []), includeProcedures: true, includeCallGraph: true, includeReferences: true);
        Assert.False(result.HasVbaProject);
        Assert.Equal(0, result.Summary.ModuleCount);
        Assert.Null(result.Modules);  // not present when no project
        Assert.Null(result.CallGraph);
        Assert.Null(result.References);
    }

    [Fact]
    public void Summary_counts_match_collections()
    {
        var p = Project(("Util", "standardModule", "Sub Log()\nEnd Sub\nSub Warn()\nLog\nEnd Sub"));
        var r = VbaSourceAnalyzer.Analyze(p, true, true, true);
        Assert.Equal(1, r.Summary.ModuleCount);
        Assert.Equal(1, r.Summary.ParsedModuleCount);
        Assert.Equal(0, r.Summary.UnparsedModuleCount);
        Assert.Equal(2, r.Summary.ProcedureCount);
        Assert.Equal(1, r.Summary.CallEdgeCount);
    }

    [Fact]
    public void Toggles_omit_collections()
    {
        var p = Project(("M", "standardModule", "Sub A()\nEnd Sub"));
        var r = VbaSourceAnalyzer.Analyze(p, includeProcedures: false, includeCallGraph: false, includeReferences: false);
        Assert.Null(r.Modules);
        Assert.Null(r.CallGraph);
        Assert.Null(r.References);
        // Summary still populated — analysis runs internally to compute counts.
        Assert.Equal(1, r.Summary.ProcedureCount);
    }

    [Fact]
    public void Module_too_large_marked_unparsed()
    {
        var bigSource = string.Join("\n", Enumerable.Repeat("x = 1", 5001));
        var p = Project(("Big", "standardModule", "Sub A()\n" + bigSource + "\nEnd Sub"));
        var r = VbaSourceAnalyzer.Analyze(p, true, false, false);
        Assert.Single(r.Modules!);
        Assert.False(r.Modules![0].Parsed);
        Assert.Equal("module_too_large", r.Modules[0].Reason);
        Assert.Equal(1, r.Summary.UnparsedModuleCount);
    }

    [Fact]
    public void Event_handler_count_in_summary()
    {
        var p = Project(("ThisWorkbook", "documentModule",
            "Private Sub Workbook_Open()\nEnd Sub"));
        var r = VbaSourceAnalyzer.Analyze(p, true, false, false);
        Assert.Equal(1, r.Summary.EventHandlerCount);
    }
}
