using System.Diagnostics;
using McpOffice.Services.Excel;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class AirSampleRenderTests
{
    private const string SamplePath = @"C:\Projects\mcpOffice-samples\Air.xlsm";

    [Fact]
    public void Whole_workbook_render_throws_graph_too_large()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var act = () => svc.RenderVbaCallgraph(
            SamplePath,
            format: "mermaid",
            moduleName: null,
            procedureName: null,
            depth: 2,
            direction: "both",
            layout: "clustered",
            maxNodes: 300);

        var ex = Assert.Throws<McpException>(act);
        Assert.Contains("graph_too_large", ex.Message);
    }

    [Fact]
    public void Single_module_render_succeeds_under_max_nodes()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var analysis = svc.AnalyzeVba(SamplePath, true, true, false);
        var probe = analysis.Modules!.First(m => m.Parsed && m.Procedures.Count > 0);

        var output = svc.RenderVbaCallgraph(
            SamplePath,
            format: "mermaid",
            moduleName: probe.Name,
            procedureName: null,
            depth: 2,
            direction: "both",
            layout: "clustered",
            maxNodes: 300);

        Assert.StartsWith("flowchart TD", output);
        Assert.Contains("subgraph", output);
    }

    [Fact]
    public void Pipeline_completes_under_500ms()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var analysis = svc.AnalyzeVba(SamplePath, true, true, false);
        var probe = analysis.Modules!.First(m => m.Parsed && m.Procedures.Count > 0);

        var sw = Stopwatch.StartNew();
        _ = svc.RenderVbaCallgraph(
            SamplePath,
            format: "mermaid",
            moduleName: probe.Name,
            procedureName: probe.Procedures[0].Name,
            depth: 1,
            direction: "both",
            layout: "clustered",
            maxNodes: 300);
        sw.Stop();

        Assert.True(sw.ElapsedMilliseconds < 500,
            $"render pipeline took {sw.ElapsedMilliseconds}ms, expected < 500ms");
    }
}
