using McpOffice.Models;
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

public class AirSampleAnalysisTests
{
    private const string SamplePath = @"C:\Projects\mcpOffice-samples\Air.xlsm";

    [Fact]
    public void Analyzes_real_air_workbook_without_exceptions()
    {
        if (!File.Exists(SamplePath)) return;  // gracefully no-op on machines without the sample

        var svc = new ExcelWorkbookService();
        var analysis = svc.AnalyzeVba(SamplePath,
            includeProcedures: true, includeCallGraph: true, includeReferences: true);

        Assert.True(analysis.HasVbaProject);
        Assert.NotNull(analysis.Modules);
        Assert.NotNull(analysis.CallGraph);
        Assert.NotNull(analysis.References);

        // Every module is either parsed or carries a reason — never both null.
        foreach (var m in analysis.Modules!)
            Assert.True(m.Parsed || m.Reason is not null);

        // Plausible floors for 107 modules of real macro code.
        Assert.True(analysis.Summary.ProcedureCount > 50,
            $"expected > 50 procedures, got {analysis.Summary.ProcedureCount}");
        Assert.NotEmpty(analysis.CallGraph!);
        Assert.Contains(analysis.References!.ObjectModel, r => r.Api == "Worksheets");
        Assert.Contains(analysis.References!.ObjectModel, r => r.Api == "Range");
    }
}
