using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

public class SyntheticConversionHintsTests
{
    [Fact]
    public void Suggests_conversion_for_synthetic_workbook()
    {
        var path = TestFixtures.Path("synthetic-vba.xlsm");
        var svc = new ExcelWorkbookService();

        var result = svc.SuggestVbaConversion(path, moduleName: null, targetParadigm: null);

        Assert.True(result.Summary.TotalProcedures > 0);
        Assert.Equal(result.Summary.TotalProcedures, result.Summary.HintedProcedures);
        Assert.Equal(4, result.Summary.ModuleCount);                          // synthetic-vba has 4 modules
        Assert.NotEmpty(result.ProcedureHints);
        Assert.NotEmpty(result.ModuleCoupling);

        // Every hint has a non-null axes object and rationale.
        Assert.All(result.ProcedureHints, h =>
        {
            Assert.NotNull(h.Axes);
            Assert.False(string.IsNullOrWhiteSpace(h.Rationale));
            Assert.Null(h.CsharpSuggestion);  // no paradigm requested
        });
    }

    [Fact]
    public void Suggests_conversion_with_classLibrary_paradigm_populates_csharpSuggestion()
    {
        var path = TestFixtures.Path("synthetic-vba.xlsm");
        var svc = new ExcelWorkbookService();

        var result = svc.SuggestVbaConversion(path, moduleName: null, targetParadigm: "classLibrary");

        Assert.Equal("classLibrary", result.Summary.TargetParadigm);
        Assert.All(result.ProcedureHints, h => Assert.NotNull(h.CsharpSuggestion));
    }

    [Fact]
    public void Suggests_conversion_filtered_by_module_name()
    {
        var path = TestFixtures.Path("synthetic-vba.xlsm");
        var svc = new ExcelWorkbookService();

        // Discover a module name first via the analyzer.
        var analysis = svc.AnalyzeVba(path, includeProcedures: true, includeCallGraph: false, includeReferences: false);
        var firstModule = analysis.Modules!.First(m => m.Procedures.Count > 0).Name;

        var result = svc.SuggestVbaConversion(path, moduleName: firstModule, targetParadigm: null);

        Assert.True(result.Summary.HintedProcedures < result.Summary.TotalProcedures
                    || result.Summary.ModuleCount == 1);
        Assert.All(result.ProcedureHints, h => Assert.Equal(firstModule, h.Module));
        Assert.True(result.ModuleCoupling.Count >= 1);   // whole-workbook coupling unaffected by filter
    }
}
