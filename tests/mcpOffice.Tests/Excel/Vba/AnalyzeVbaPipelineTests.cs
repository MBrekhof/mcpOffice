using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

/// <summary>
/// Exercises the full OLE extraction → decompression → analysis pipeline using
/// synthetic vbaProject.bin blobs built by VbaProjectBinBuilder. Verifies that
/// VbaProjectReader.ReadVbaProjectBin + VbaSourceAnalyzer.Analyze compose correctly
/// without relying on real .xlsm sample files.
/// </summary>
public class AnalyzeVbaPipelineTests
{
    [Fact]
    public void Full_pipeline_detects_procedures_and_cross_module_call_edge()
    {
        // Two modules: Caller has Sub Run() that calls DoLog, Utils has Sub DoLog().
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("Caller", "Caller", "Sub Run()\r\nDoLog\r\nEnd Sub"),
            new ModuleSpec("Utils",  "Utils",  "Sub DoLog()\r\nEnd Sub")
        ]);

        using var stream = new MemoryStream(bytes);
        var project = new VbaProjectReader().ReadVbaProjectBin(stream, "synthetic");

        Assert.True(project.HasVbaProject);
        Assert.Equal(2, project.Modules.Count);

        var analysis = VbaSourceAnalyzer.Analyze(
            project,
            includeProcedures: true,
            includeCallGraph: true,
            includeReferences: false);

        // Summary checks
        Assert.True(analysis.HasVbaProject);
        Assert.Equal(2, analysis.Summary.ModuleCount);
        Assert.Equal(2, analysis.Summary.ParsedModuleCount);
        Assert.True(analysis.Summary.ProcedureCount >= 2, "Expected at least 2 procedures");
        Assert.True(analysis.Summary.CallEdgeCount >= 1, "Expected at least one call edge");

        // At least one procedure per module
        Assert.NotNull(analysis.Modules);
        Assert.All(analysis.Modules!, m => Assert.NotEmpty(m.Procedures));

        // Cross-module call edge: Caller.Run -> Utils.DoLog
        Assert.NotNull(analysis.CallGraph);
        var edge = Assert.Single(analysis.CallGraph!,
            e => e.From == "Caller.Run" && e.To == "Utils.DoLog");
        Assert.True(edge.Resolved);
    }

    [Fact]
    public void Pipeline_handles_empty_module_gracefully()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("Module1", "Module1", "Sub Hello()\r\nEnd Sub"),
            new ModuleSpec("Empty",   "Empty",   "")
        ]);

        using var stream = new MemoryStream(bytes);
        var project = new VbaProjectReader().ReadVbaProjectBin(stream, "synthetic-empty");
        var analysis = VbaSourceAnalyzer.Analyze(
            project,
            includeProcedures: true,
            includeCallGraph: false,
            includeReferences: false);

        Assert.True(analysis.HasVbaProject);
        Assert.Equal(2, analysis.Summary.ModuleCount);
        Assert.Equal(1, analysis.Summary.ParsedModuleCount);   // Empty module is unparsed
        Assert.Equal(1, analysis.Summary.UnparsedModuleCount);
        Assert.Equal(1, analysis.Summary.ProcedureCount);
    }
}
