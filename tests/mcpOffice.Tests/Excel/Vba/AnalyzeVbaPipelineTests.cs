using McpOffice.Models;
using McpOffice.Services.Excel.Vba;
using ModelContextProtocol;

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
    public void ModuleName_filter_narrows_modules_callgraph_and_references_to_focal_module()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("Caller", "Caller", "Sub Run()\r\nDoLog\r\nWorksheets(\"Foo\").Range(\"A1\").Value = 1\r\nEnd Sub"),
            new ModuleSpec("Utils",  "Utils",  "Sub DoLog()\r\nWorksheets(\"Bar\").Range(\"B2\").Value = 2\r\nEnd Sub"),
            new ModuleSpec("Other",  "Other",  "Sub Solo()\r\nWorksheets(\"Baz\").Range(\"C3\").Value = 3\r\nEnd Sub"),
        ]);

        using var stream = new MemoryStream(bytes);
        var project = new VbaProjectReader().ReadVbaProjectBin(stream, "synthetic-filter");

        var analysis = VbaSourceAnalyzer.Analyze(
            project,
            includeProcedures: true,
            includeCallGraph: true,
            includeReferences: true,
            moduleName: "Caller");

        // Summary still describes the whole workbook (3 modules, 3 procs, all the edges/refs).
        Assert.Equal(3, analysis.Summary.ModuleCount);
        Assert.Equal(3, analysis.Summary.ProcedureCount);

        // Modules array filtered to just Caller.
        Assert.NotNull(analysis.Modules);
        var module = Assert.Single(analysis.Modules!);
        Assert.Equal("Caller", module.Name);

        // Call graph: edges involving Caller (caller side or resolved-callee side).
        Assert.NotNull(analysis.CallGraph);
        Assert.All(analysis.CallGraph!, e =>
            Assert.True(
                e.Site.Module == "Caller" || (e.Resolved && e.To.StartsWith("Caller.", StringComparison.Ordinal)),
                $"Edge {e.From}->{e.To} (site={e.Site.Module}) leaked through filter"));
        Assert.Contains(analysis.CallGraph!, e => e.From == "Caller.Run" && e.To == "Utils.DoLog");
        Assert.DoesNotContain(analysis.CallGraph!, e => e.From == "Other.Solo");

        // References: only entries from Caller.
        Assert.NotNull(analysis.References);
        Assert.All(analysis.References!.ObjectModel, r => Assert.Equal("Caller", r.Module));
        Assert.All(analysis.References!.Dependencies, d => Assert.Equal("Caller", d.Module));
        Assert.Contains(analysis.References!.ObjectModel, r => r.Literal == "Foo");
        Assert.DoesNotContain(analysis.References!.ObjectModel, r => r.Literal == "Bar");
        Assert.DoesNotContain(analysis.References!.ObjectModel, r => r.Literal == "Baz");
    }

    [Fact]
    public void ModuleName_filter_is_case_insensitive()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("Module1", "Module1", "Sub Hello()\r\nEnd Sub"),
        ]);
        using var stream = new MemoryStream(bytes);
        var project = new VbaProjectReader().ReadVbaProjectBin(stream, "synthetic-case");

        var analysis = VbaSourceAnalyzer.Analyze(
            project,
            includeProcedures: true,
            includeCallGraph: false,
            includeReferences: false,
            moduleName: "MODULE1");

        Assert.NotNull(analysis.Modules);
        var module = Assert.Single(analysis.Modules!);
        Assert.Equal("Module1", module.Name); // original casing preserved
    }

    [Fact]
    public void ModuleName_filter_throws_module_not_found_when_unknown()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("Module1", "Module1", "Sub Hello()\r\nEnd Sub"),
            new ModuleSpec("Module2", "Module2", "Sub World()\r\nEnd Sub"),
        ]);
        using var stream = new MemoryStream(bytes);
        var project = new VbaProjectReader().ReadVbaProjectBin(stream, "synthetic-missing");

        var ex = Assert.Throws<McpException>(() =>
            VbaSourceAnalyzer.Analyze(
                project,
                includeProcedures: true,
                includeCallGraph: false,
                includeReferences: false,
                moduleName: "DoesNotExist"));

        Assert.Contains("module_not_found", ex.Message);
        Assert.Contains("DoesNotExist", ex.Message);
        Assert.Contains("Module1", ex.Message);
        Assert.Contains("Module2", ex.Message);
    }

    [Fact]
    public void ModuleName_null_or_empty_preserves_full_output()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("A", "A", "Sub Foo()\r\nEnd Sub"),
            new ModuleSpec("B", "B", "Sub Bar()\r\nEnd Sub"),
        ]);
        using var stream = new MemoryStream(bytes);
        var project = new VbaProjectReader().ReadVbaProjectBin(stream, "synthetic-nullfilter");

        var withNull = VbaSourceAnalyzer.Analyze(project, true, true, true, moduleName: null);
        var withEmpty = VbaSourceAnalyzer.Analyze(project, true, true, true, moduleName: "");
        var withWhitespace = VbaSourceAnalyzer.Analyze(project, true, true, true, moduleName: "   ");

        foreach (var a in new[] { withNull, withEmpty, withWhitespace })
        {
            Assert.NotNull(a.Modules);
            Assert.Equal(2, a.Modules!.Count);
        }
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
