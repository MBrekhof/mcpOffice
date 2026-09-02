using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

/// <summary>Gated real-world checks for excel_map_vba_sheet_access and excel_compare_vba_corpus.</summary>
public class VbaSheetAccessSampleTests
{
    private const string Samples = @"C:\Projects\mcpOffice-samples";
    private const string Ring = Samples + @"\RingOnderzoek.xlsm";
    private const string Air  = Samples + @"\Air.xlsm";

    [Fact]
    public void Synthetic_fixture_maps_without_error()
    {
        var r = new ExcelWorkbookService().MapVbaSheetAccess(TestFixtures.Path("synthetic-vba.xlsm"), null, null, true);
        Assert.True(r.HasVbaProject);
        Assert.Equal(r.Summary.SiteCount, r.Summary.ResolvedCount + r.Summary.UnresolvedCount);
    }

    [Fact]
    public void RingOnderzoek_resolves_dutch_codenames_to_sheet_names()
    {
        if (!File.Exists(Ring)) return;
        var r = new ExcelWorkbookService().MapVbaSheetAccess(Ring, null, null, true);

        Assert.True(r.HasVbaProject);
        Assert.True(r.Summary.ResolvedCount > 0, "expected resolved sheet access");
        Assert.All(r.Sheets, s => Assert.False(string.IsNullOrEmpty(s.Name)));
        Assert.Contains(r.Sheets, s => s.CodeName is not null && s.CodeName.StartsWith("Blad"));
    }

    [Fact]
    public void Air_maps_and_stays_within_caps()
    {
        if (!File.Exists(Air)) return;
        var r = new ExcelWorkbookService().MapVbaSheetAccess(Air, null, null, true);
        Assert.True(r.HasVbaProject);
        Assert.True(r.Summary.SiteCount > 100, "Air has thousands of object-model sites");
        Assert.True(r.SheetAccess.Count <= 1000);
    }

    [Fact]
    public void Corpus_over_the_samples_directory_finds_shared_code()
    {
        if (!Directory.Exists(Samples)) return;
        var r = new ExcelWorkbookService().CompareVbaCorpus(null, Samples, minOccurrences: 2, maxProcedures: 50, includeNearDuplicates: true);

        Assert.True(r.Summary.WorkbookCount >= 2);
        Assert.All(r.SharedProcedures, g => Assert.True(g.Occurrences.Select(o => o.Workbook).Distinct().Count() >= 2));
        Assert.True(r.SharedProcedures.Count <= 50);
    }
}
