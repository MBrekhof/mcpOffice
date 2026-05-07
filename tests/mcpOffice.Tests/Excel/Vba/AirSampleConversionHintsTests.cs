using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

public class AirSampleConversionHintsTests
{
    private const string SamplePath = @"C:\Projects\mcpOffice-samples\Air.xlsm";

    [Fact]
    public void Suggests_conversion_for_air_workbook_within_budget()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var result = svc.SuggestVbaConversion(SamplePath, moduleName: null, targetParadigm: null);
        sw.Stop();

        // Plausible floors against the documented Air.xlsm metrics
        // (107 modules, 200 procedures).
        Assert.True(result.Summary.ModuleCount > 50,
            $"expected > 50 modules, got {result.Summary.ModuleCount}");
        Assert.True(result.Summary.TotalProcedures > 100,
            $"expected > 100 procedures, got {result.Summary.TotalProcedures}");
        Assert.True(result.Summary.TotalProcedures == result.Summary.HintedProcedures);

        // Every procedure has a hint.
        Assert.Equal(result.Summary.TotalProcedures, result.ProcedureHints.Count);

        // Coupling block populated.
        Assert.True(result.ModuleCoupling.Count >= result.Summary.ModuleCount);
        Assert.NotEmpty(result.CouplingPairs);

        // Performance budget — generous to absorb cold-cache variance.
        Assert.True(sw.ElapsedMilliseconds < 600,
            $"expected < 600 ms, got {sw.ElapsedMilliseconds} ms");
    }

    [Fact]
    public void ClassLibrary_paradigm_produces_static_methods_and_manual_reviews()
    {
        if (!File.Exists(SamplePath)) return;

        var svc = new ExcelWorkbookService();
        var result = svc.SuggestVbaConversion(SamplePath, null, "classLibrary");

        Assert.Contains(result.ProcedureHints,
            h => h.CsharpSuggestion?.TargetType == "staticMethod");
        Assert.Contains(result.ProcedureHints,
            h => h.CsharpSuggestion?.TargetType == "requiresManualReview");
    }
}
