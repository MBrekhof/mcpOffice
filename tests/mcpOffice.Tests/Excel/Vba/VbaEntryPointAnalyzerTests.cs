using McpOffice.Models;
using McpOffice.Services.Excel.Vba;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class VbaEntryPointAnalyzerTests
{
    private const string Module1 = """
        Public Sub GetILIS()
            Helper
        End Sub
        Private Sub Helper()
        End Sub
        Private Sub Orphan()
        End Sub
        Public Function Score(x As Double) As Double
            Score = x * 2
        End Function
        Public Sub Auto_Open()
            Application.OnTime Now, "Later"
        End Sub
        Public Sub Later()
        End Sub
        Public Sub NextDate()
        End Sub
        """;

    private const string Blad1 = """
        Private Sub Worksheet_Change(ByVal Target As Range)
        End Sub
        """;

    private const string DrawingXml = """
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xdr:twoCellAnchor>
            <xdr:pic macro="[0]!GetILIS"><xdr:nvPicPr><xdr:cNvPr id="86" name="Picture 84"/></xdr:nvPicPr></xdr:pic>
          </xdr:twoCellAnchor>
          <xdr:twoCellAnchor>
            <xdr:sp macro="'Copy_results(2)'"><xdr:nvSpPr><xdr:cNvPr id="3" name="Rectangle 2"/></xdr:nvSpPr></xdr:sp>
          </xdr:twoCellAnchor>
          <xdr:twoCellAnchor>
            <xdr:sp macro=""><xdr:nvSpPr><xdr:cNvPr id="4" name="Button 14"/></xdr:nvSpPr></xdr:sp>
          </xdr:twoCellAnchor>
        </xdr:wsDr>
        """;

    private const string Vml = """
        <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
          <v:shape id="_x0000_s1025" type="#_x0000_t201">
            <x:ClientData ObjectType="Button"><x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor><x:FmlaMacro>[0]!NextDate</x:FmlaMacro><x:TextHAlign>Center</x:TextHAlign></x:ClientData>
          </v:shape>
        </xml>
        """;

    private static ExcelVbaProject Project() => new(true,
    [
        new ExcelVbaModule("Module1", "standardModule", Module1.Split('\n').Length, Module1),
        new ExcelVbaModule("Blad1", "documentModule", Blad1.Split('\n').Length, Blad1),
    ]);

    private static IReadOnlyList<VbaEntryPointAnalyzer.SheetInput> Sheets() =>
    [
        new("Data", "Blad1", DrawingXml, Vml, [("B2", "Score(A2)"), ("C3", "SUM(A1:A3)"), ("D4", "Sheet2.Score(1)")]),
    ];

    private static ExcelVbaEntryPointsResult Run(string? moduleName = null, bool includeUnreachable = true) =>
        VbaEntryPointAnalyzer.Analyze(@"C:\t.xlsm", Project(), Sheets(), includeUnreachable, moduleName);

    [Fact]
    public void Shape_macro_resolves_to_procedure_with_sheet_and_shape_name()
    {
        var e = Assert.Single(Run().EntryPoints, e => e.Kind == "shapeMacro" && e.Resolved);
        Assert.Equal("Module1.GetILIS", e.Procedure);
        Assert.Equal("Data", e.Sheet);
        Assert.Equal("Picture 84", e.ShapeName);
        Assert.Equal("[0]!GetILIS", e.Target);
    }

    [Fact]
    public void Shape_macro_naming_an_unknown_procedure_is_kept_verbatim_and_counted()
    {
        // 'Copy_results(2)' parses to Copy_results (macro with an argument) but no such procedure exists here.
        var r = Run();
        var e = Assert.Single(r.EntryPoints, e => e.Kind == "shapeMacro" && !e.Resolved);
        Assert.Null(e.Procedure);
        Assert.Equal("'Copy_results(2)'", e.Target);
        Assert.Equal(1, r.Summary.UnresolvedMacroReferences);
        Assert.DoesNotContain(r.EntryPoints, e => e.ShapeName == "Button 14");   // empty macro="" is not an entry
    }

    [Fact]
    public void Form_control_macro_from_vml_resolves()
    {
        var e = Assert.Single(Run().EntryPoints, e => e.Kind == "formControlMacro");
        Assert.Equal("Module1.NextDate", e.Procedure);
        Assert.Equal("_x0000_s1025", e.ShapeName);
    }

    [Fact]
    public void Public_function_used_in_a_formula_is_a_worksheet_function()
    {
        var e = Assert.Single(Run().EntryPoints, e => e.Kind == "worksheetFunction");
        Assert.Equal("Module1.Score", e.Procedure);
        Assert.Equal(["Data!B2"], e.FormulaCells);   // "Sheet2.Score(" is qualified, not a UDF call
    }

    [Fact]
    public void Formula_calling_a_function_twice_lists_the_cell_once()
    {
        // Air.xlsm campy!K13 calls MPNindex three times in one formula and showed up three times.
        IReadOnlyList<VbaEntryPointAnalyzer.SheetInput> sheets =
            [new("Data", "Blad1", DrawingXml, Vml, [("K13", "IF(A1>0,Score(A1),Score(A2))+Score(A3)"), ("L13", "Score(B1)")])];
        var r = VbaEntryPointAnalyzer.Analyze(@"C:\t.xlsm", Project(), sheets, true, null);
        var e = Assert.Single(r.EntryPoints, e => e.Kind == "worksheetFunction");
        Assert.Equal(["Data!K13", "Data!L13"], e.FormulaCells);
    }

    [Fact]
    public void Auto_Open_and_event_handlers_are_entry_points()
    {
        var r = Run();
        Assert.Contains(r.EntryPoints, e => e.Kind == "autoMacro" && e.Procedure == "Module1.Auto_Open");
        Assert.Contains(r.EntryPoints, e => e.Kind == "eventHandler" && e.Procedure == "Blad1.Worksheet_Change");
    }

    [Fact]
    public void OnTime_literal_target_is_dynamic_dispatch_with_site_and_rescues_reachability()
    {
        var r = Run();
        var e = Assert.Single(r.EntryPoints, e => e.Kind == "dynamicDispatch");
        Assert.Equal("Module1.Later", e.Procedure);
        Assert.Equal("Later", e.Target);
        Assert.NotNull(e.Site);
        Assert.Equal("Auto_Open", e.Site!.Procedure);
        Assert.DoesNotContain(r.Unreachable!, u => u.Procedure == "Module1.Later");
    }

    [Fact]
    public void Only_the_orphan_is_unreachable_with_high_confidence()
    {
        var r = Run();
        var u = Assert.Single(r.Unreachable!);
        Assert.Equal("Module1.Orphan", u.Procedure);
        Assert.Equal("high", u.Confidence);
        Assert.Equal("Private", u.Scope);
        Assert.Equal(r.Summary.ProcedureCount - 1, r.Summary.ReachableCount);
        Assert.Equal(1, r.Summary.UnreachableCount);
    }

    [Fact]
    public void Summary_by_kind_counts_every_kind()
    {
        var byKind = Run().Summary.ByKind;
        Assert.Equal(2, byKind["shapeMacro"]);
        Assert.Equal(1, byKind["formControlMacro"]);
        Assert.Equal(1, byKind["worksheetFunction"]);
        Assert.Equal(1, byKind["autoMacro"]);
        Assert.Equal(1, byKind["eventHandler"]);
        Assert.Equal(1, byKind["dynamicDispatch"]);
    }

    [Fact]
    public void ModuleName_scopes_arrays_but_not_summary()
    {
        var r = Run(moduleName: "blad1");
        Assert.All(r.EntryPoints, e => Assert.StartsWith("Blad1.", e.Procedure));
        Assert.Empty(r.Unreachable!);
        Assert.Equal(1, r.Summary.UnreachableCount);
    }

    [Fact]
    public void Unknown_module_throws_module_not_found()
    {
        var ex = Assert.Throws<McpException>(() => Run(moduleName: "Nope"));
        Assert.Contains("module_not_found", ex.Message);
    }

    [Fact]
    public void IncludeUnreachable_false_omits_the_array()
    {
        Assert.Null(Run(includeUnreachable: false).Unreachable);
    }

    [Fact]
    public void Workbook_without_vba_project_returns_empty_result()
    {
        var r = VbaEntryPointAnalyzer.Analyze(@"C:\t.xlsx", new ExcelVbaProject(false, []), [], true, null);
        Assert.False(r.HasVbaProject);
        Assert.Empty(r.EntryPoints);
        Assert.Equal(0, r.Summary.ProcedureCount);
    }
}
