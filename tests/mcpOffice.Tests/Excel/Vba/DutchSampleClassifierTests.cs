using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

/// <summary>
/// Gated regression tests for VbaProjectReader.ClassifyKind against real Dutch-locale
/// .xlsm samples. Without OOXML-derived codenames, sheet-bound document modules in
/// Dutch Excel land as "classModule" because the legacy heuristic only matches the
/// English "Sheet*" prefix. Tests skip when the sample files are absent.
/// </summary>
public class DutchSampleClassifierTests
{
    private const string RingOnderzoek = @"C:\Projects\mcpOffice-samples\RingOnderzoek.xlsm";
    private const string Balans = @"C:\Projects\mcpOffice-samples\Balans.xlsm";

    [Fact]
    public void RingOnderzoek_Blad_modules_are_documentModule()
    {
        if (!File.Exists(RingOnderzoek)) return;

        var svc = new ExcelWorkbookService();
        var analysis = svc.AnalyzeVba(RingOnderzoek,
            includeProcedures: true, includeCallGraph: false, includeReferences: false);

        Assert.NotNull(analysis.Modules);
        var blad1 = analysis.Modules!.SingleOrDefault(m => m.Name == "Blad1");
        var blad3 = analysis.Modules!.SingleOrDefault(m => m.Name == "Blad3");
        var thisWb = analysis.Modules!.SingleOrDefault(m => m.Name == "ThisWorkbook");

        Assert.NotNull(blad1);
        Assert.NotNull(blad3);
        Assert.NotNull(thisWb);

        Assert.Equal("documentModule", blad1!.Kind);
        Assert.Equal("documentModule", blad3!.Kind);
        Assert.Equal("documentModule", thisWb!.Kind);
    }

    [Fact]
    public void Balans_Blad3_is_documentModule()
    {
        if (!File.Exists(Balans)) return;

        var svc = new ExcelWorkbookService();
        var analysis = svc.AnalyzeVba(Balans,
            includeProcedures: true, includeCallGraph: false, includeReferences: false);

        Assert.NotNull(analysis.Modules);
        var blad3 = analysis.Modules!.SingleOrDefault(m => m.Name == "Blad3");
        Assert.NotNull(blad3);
        Assert.Equal("documentModule", blad3!.Kind);
    }
}
