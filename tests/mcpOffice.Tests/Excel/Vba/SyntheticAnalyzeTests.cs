using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

/// <summary>
/// End-to-end test that exercises the full pipeline (xlsm → vbaProject.bin →
/// MS-OVBA decompression → dir parse → module sources → analyzer) against a
/// real Excel-authored fixture committed to the repo. Mirrors the gated
/// AirSampleAnalysisTests but runs unconditionally — `synthetic-vba.xlsm` is
/// in tests/fixtures/ and ships with the suite.
///
/// Regenerate the fixture with tests/fixtures/Generate-SyntheticVbaXlsm.ps1
/// (requires Excel + "Trust access to the VBA project object model").
/// </summary>
public class SyntheticAnalyzeTests
{
    [Fact]
    public void Analyzes_synthetic_workbook_full_pipeline()
    {
        var path = TestFixtures.Path("synthetic-vba.xlsm");
        var svc = new ExcelWorkbookService();

        var analysis = svc.AnalyzeVba(path,
            includeProcedures: true, includeCallGraph: true, includeReferences: true);

        Assert.True(analysis.HasVbaProject);
        Assert.NotNull(analysis.Modules);
        Assert.NotNull(analysis.CallGraph);
        Assert.NotNull(analysis.References);

        // 4 modules: ThisWorkbook + sheet doc-module + Module1 + Class1.
        Assert.Equal(4, analysis.Summary.ModuleCount);

        var module1 = Assert.Single(analysis.Modules!, m => m.Name == "Module1");
        Assert.Equal("standardModule", module1.Kind);
        Assert.True(module1.Parsed);
        Assert.Equal(["Main", "Process", "Variadic", "StatefulCount"],
            module1.Procedures.Select(p => p.Name).ToArray());

        // Static Sub StatefulCount: prove the (Static ...) prefix is consumed by the scanner.
        var stateful = module1.Procedures.Single(p => p.Name == "StatefulCount");
        Assert.Equal("Sub", stateful.Kind);

        // ParamArray on Variadic: parameter survives the modifier strip, type captured.
        var variadic = module1.Procedures.Single(p => p.Name == "Variadic");
        var arg = Assert.Single(variadic.Parameters);
        Assert.Equal("Variant", arg.Type);

        // Class1: classModule, single Public Sub Greet(who As String).
        var class1 = Assert.Single(analysis.Modules!, m => m.Name == "Class1");
        Assert.Equal("classModule", class1.Kind);
        var greet = Assert.Single(class1.Procedures);
        Assert.Equal("Greet", greet.Name);
        Assert.Equal("Public", greet.Scope);

        // Document modules: workbook codename is always "ThisWorkbook" (VBA built-in).
        // The sheet codename is locale-dependent (Dutch: "Blad1"). Match on Workbook_Open
        // / Worksheet_Change rather than on module name.
        var docModules = analysis.Modules!.Where(m => m.Kind == "documentModule").ToList();
        Assert.Equal(2, docModules.Count);

        Assert.Contains(docModules, m =>
            m.Procedures.Any(p => p.Name == "Workbook_Open" && p.IsEventHandler && p.EventTarget == "Workbook"));
        Assert.Contains(docModules, m =>
            m.Procedures.Any(p => p.Name == "Worksheet_Change" && p.IsEventHandler && p.EventTarget == "Worksheet"));

        // Cross-module edge from ThisWorkbook.Workbook_Open -> Module1.Main: load-bearing,
        // proves bareword resolution + cross-module FQN composition. Note: Main's call to
        // Process is intentionally not asserted — VbaCallGraphBuilder's regex requires the
        // callee to be followed by `(`, `.`, or end-of-line, so bareword-with-args (`Process arg`)
        // doesn't surface as an edge by design.
        Assert.Contains(analysis.CallGraph!,
            e => e.From == "ThisWorkbook.Workbook_Open"
                 && e.To == "Module1.Main"
                 && e.Resolved);

        // Object-model references from Main: Worksheets("Data").Range("A1").
        Assert.Contains(analysis.References!.ObjectModel,
            r => r.Module == "Module1" && r.Api == "Worksheets" && r.Literal == "Data");
        Assert.Contains(analysis.References!.ObjectModel,
            r => r.Module == "Module1" && r.Api == "Range" && r.Literal == "A1");
    }
}
