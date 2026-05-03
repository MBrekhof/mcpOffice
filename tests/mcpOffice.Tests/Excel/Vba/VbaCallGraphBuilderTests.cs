using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaCallGraphBuilderTests
{
    private static IReadOnlyList<ExcelVbaCallEdge> Build(params (string moduleName, string moduleKind, string source)[] modules)
    {
        var scanned = new List<(string ModuleName, IReadOnlyList<CleanedLine> Lines, IReadOnlyList<ScannedProcedure> Procs)>();
        foreach (var (n, k, s) in modules)
        {
            var lines = VbaLineCleaner.Clean(s);
            var procs = VbaProcedureScanner.Scan(k, n, lines);
            scanned.Add((n, lines, procs));
        }
        return VbaCallGraphBuilder.Build(scanned);
    }

    [Fact]
    public void Resolves_direct_call_within_module()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nB\nEnd Sub\nSub B()\nEnd Sub"));
        var edge = Assert.Single(edges);
        Assert.Equal("M.A", edge.From);
        Assert.Equal("M.B", edge.To);
        Assert.True(edge.Resolved);
    }

    [Fact]
    public void Resolves_call_keyword_form()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nCall B\nEnd Sub\nSub B()\nEnd Sub"));
        Assert.Single(edges);
        Assert.True(edges[0].Resolved);
    }

    [Fact]
    public void Resolves_cross_module_call()
    {
        var edges = Build(
            ("Caller", "standardModule", "Sub A()\nDoLog\nEnd Sub"),
            ("Utils", "standardModule", "Sub DoLog()\nEnd Sub"));
        var edge = Assert.Single(edges);
        Assert.Equal("Caller.A", edge.From);
        Assert.Equal("Utils.DoLog", edge.To);
        Assert.True(edge.Resolved);
    }

    [Fact]
    public void Captures_application_run_as_dynamic_unresolved()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nApplication.Run \"OtherWb.xlsm!Foo\"\nEnd Sub"));
        var edge = Assert.Single(edges);
        Assert.False(edge.Resolved);
        Assert.Equal("OtherWb.xlsm!Foo", edge.To);
    }

    [Fact]
    public void Skips_vba_keywords_and_string_sentinels()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nIf x Then\nDim y As Long\nEnd If\nEnd Sub"));
        Assert.Empty(edges);
    }

    [Fact]
    public void Records_call_site_line()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nx = 1\nB\nEnd Sub\nSub B()\nEnd Sub"));
        Assert.Equal(3, edges[0].Site.Line);
    }
}
