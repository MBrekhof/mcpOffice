using McpOffice.Models;
using McpOffice.Services.Excel.Vba;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class VbaSheetAccessAnalyzerTests
{
    private const string Module1 = """
        Sub A()
            Worksheets("Data").Range("A1").Value = 1
            v = Worksheets("Data").Range("A1")
            w = ActiveSheet.Range("B1")
        End Sub
        Sub B()
            v = Worksheets("Config").Range("Total")
        End Sub
        """;

    private static readonly IReadOnlyList<VbaSheetAccessResolver.SheetName> Sheets = [new("Data", "Blad1"), new("Config", "Blad2")];
    private static readonly IReadOnlyList<VbaSheetAccessResolver.DefinedName> Names = [new("Total", null, "=Config!$C$2")];

    private static ExcelVbaSheetAccessResult Run(string? moduleName = null, string? sheetName = null, bool includeUnresolved = true,
                                                 bool includeRecords = true, int maxRecords = 100) =>
        VbaSheetAccessAnalyzer.Analyze(@"C:\t.xlsm",
            new ExcelVbaProject(true, [new ExcelVbaModule("Module1", "standardModule", 9, Module1)]),
            Sheets, Names, moduleName, sheetName, includeUnresolved, includeRecords, maxRecords);

    [Fact]
    public void Sites_on_the_same_target_aggregate_into_one_record_with_mode_both()
    {
        var r = Run();
        var a1 = Assert.Single(r.SheetAccess, a => a.Procedure == "Module1.A" && a.Sheet?.Name == "Data");
        Assert.Equal("both", a1.Mode);
        Assert.Equal(2, a1.SiteCount);
        Assert.Equal("Blad1", a1.Sheet!.CodeName);
        Assert.Equal(("range", "A1"), (a1.Target.Kind, a1.Target.Address));
    }

    [Fact]
    public void Unresolved_sites_are_kept_with_a_reason()
    {
        var u = Assert.Single(Run().SheetAccess, a => a.Sheet is null);
        Assert.Equal("activeSheet", u.UnresolvedReason);
        Assert.Equal("Module1.A", u.Procedure);
    }

    [Fact]
    public void Defined_name_record_carries_name_and_resolved_sheet()
    {
        var t = Assert.Single(Run().SheetAccess, a => a.Procedure == "Module1.B");
        Assert.Equal("definedName", t.Target.Kind);
        Assert.Equal("Total", t.Target.DefinedName);
        Assert.Equal("Config", t.Sheet!.Name);
    }

    [Fact]
    public void Sheets_rollup_lists_readers_and_writers()
    {
        var data = Assert.Single(Run().Sheets, s => s.Name == "Data");
        Assert.Equal(["Module1.A"], data.Readers);
        Assert.Equal(["Module1.A"], data.Writers);
        Assert.Equal((1, 1), (data.ReadSites, data.WriteSites));
        Assert.Equal(2, Run().Sheets.Count);
    }

    [Fact]
    public void Summary_counts_sites_sheets_and_procedures()
    {
        var s = Run().Summary;
        Assert.Equal((4, 3, 1, 2, 2), (s.SiteCount, s.ResolvedCount, s.UnresolvedCount, s.SheetCount, s.ProcedureCount));
    }

    [Fact]
    public void Filters_scope_records_but_not_summary()
    {
        Assert.All(Run(moduleName: "module1").SheetAccess, a => Assert.StartsWith("Module1.", a.Procedure));
        var bySheet = Run(sheetName: "config");
        Assert.All(bySheet.SheetAccess, a => Assert.Equal("Config", a.Sheet!.Name));
        Assert.Single(bySheet.Sheets);
        Assert.Equal(4, bySheet.Summary.SiteCount);
        Assert.DoesNotContain(Run(includeUnresolved: false).SheetAccess, a => a.Sheet is null);
    }

    [Fact]
    public void Unknown_module_or_sheet_throws()
    {
        Assert.Contains("module_not_found", Assert.Throws<McpException>(() => Run(moduleName: "Nope")).Message);
        Assert.Contains("sheet_not_found", Assert.Throws<McpException>(() => Run(sheetName: "Nope")).Message);
    }

    [Fact]
    public void Include_records_false_returns_summary_and_rollup_only()
    {
        var r = Run(includeRecords: false);
        Assert.Empty(r.SheetAccess);
        Assert.False(r.Truncated);
        Assert.Equal(["Config", "Data"], r.Sheets.Select(s => s.Name));
        Assert.Equal(Run().Summary, r.Summary);
    }

    [Fact]
    public void Max_records_caps_the_records_and_flags_truncation_but_not_the_rollup()
    {
        var r = Run(maxRecords: 1);
        Assert.Single(r.SheetAccess);
        Assert.True(r.Truncated);
        Assert.Equal(2, r.Sheets.Count);
    }

    [Fact]
    public void Workbook_without_vba_project_is_empty()
    {
        var r = VbaSheetAccessAnalyzer.Analyze(@"C:\t.xlsx", new ExcelVbaProject(false, []), [], [], null, null, true);
        Assert.False(r.HasVbaProject);
        Assert.Empty(r.SheetAccess);
    }
}
