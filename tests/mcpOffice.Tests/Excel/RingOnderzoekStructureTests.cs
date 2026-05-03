using DevExpress.Spreadsheet;
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel;

/// <summary>
/// Gated regression + bug-watchdog tests for excel_get_structure against
/// RingOnderzoek.xlsm. The DevExpress.Spreadsheet WorksheetCollection on this
/// file is internally inconsistent: Count returns 1, foreach yields 0, and
/// Worksheets[0] throws. Our service must not throw; the watchdogs document
/// the upstream bug and will fail (signalling we can drop the workaround) if
/// DevExpress ever fixes it.
///
/// All tests are gated on the sample file being present (skipped on machines
/// where it is absent).
/// </summary>
public class RingOnderzoekStructureTests
{
    private const string SamplePath = @"C:\Projects\mcpOffice-samples\RingOnderzoek.xlsm";

    [Fact]
    public void GetStructure_does_not_throw_on_workbook_with_broken_worksheet_indexer()
    {
        if (!File.Exists(SamplePath)) return;
        var svc = new ExcelWorkbookService();
        var s = svc.GetStructure(SamplePath, includeSheets: true, includeFormulaCounts: true, includeDefinedNames: true);

        // Service degrades gracefully: returns what could be enumerated rather than throwing.
        // SheetCount matches the sheets array (internal consistency), even though
        // workbook.Worksheets.Count would report 1.
        Assert.NotNull(s.Sheets);
        Assert.Equal(s.Sheets!.Count, s.SheetCount);
        Assert.NotNull(s.DefinedNames);
        Assert.Equal(s.DefinedNames!.Count, s.DefinedNameCount);
    }

    [Fact]
    public void DevExpress_bug_watchdog_count_and_enumeration_disagree()
    {
        if (!File.Exists(SamplePath)) return;
        using var wb = new Workbook();
        wb.LoadDocument(SamplePath);

        var headerCount = wb.Worksheets.Count;
        var enumeratedCount = 0;
        foreach (var _ in wb.Worksheets) enumeratedCount++;

        Assert.Equal(1, headerCount);
        Assert.Equal(0, enumeratedCount);
    }

    [Fact]
    public void DevExpress_bug_watchdog_indexer_throws_at_zero()
    {
        if (!File.Exists(SamplePath)) return;
        using var wb = new Workbook();
        wb.LoadDocument(SamplePath);

        Assert.Throws<ArgumentOutOfRangeException>(() => _ = wb.Worksheets[0]);
    }
}
