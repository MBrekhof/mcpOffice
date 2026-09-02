using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaSheetAccessResolverTests
{
    private static readonly IReadOnlyList<VbaSheetAccessResolver.SheetName> Sheets =
    [
        new("Data", "Blad1"),
        new("Config", "Blad2"),
        new("My Sheet", "Blad3"),
    ];

    private static readonly IReadOnlyList<VbaSheetAccessResolver.DefinedName> Names =
    [
        new("Total", null, "=Config!$C$2"),
        new("Prices", null, "'My Sheet'!$A$1:$B$9"),
        new("Rate", null, "0.21"),
    ];

    private static IReadOnlyList<VbaSheetAccessResolver.AccessSite> Resolve(string body, string module = "Module1", string kind = "standardModule")
    {
        var source = $"Sub P()\n{body}\nEnd Sub";
        var lines = VbaLineCleaner.Clean(source);
        var procs = VbaProcedureScanner.Scan(kind, module, lines);
        return VbaSheetAccessResolver.Resolve(module, kind, lines, procs, Sheets, Names);
    }

    private static VbaSheetAccessResolver.AccessSite One(string body, string module = "Module1", string kind = "standardModule") =>
        Assert.Single(Resolve(body, module, kind));

    [Fact]
    public void Worksheets_literal_resolves_by_name()
    {
        var s = One("x = Worksheets(\"Data\").Range(\"A1\").Value");
        Assert.Equal("Data", s.SheetName);
        Assert.Equal("Blad1", s.CodeName);
        Assert.Equal("range", s.TargetKind);
        Assert.Equal("A1", s.Address);
        Assert.Equal("read", s.Mode);
        Assert.Null(s.UnresolvedReason);
    }

    [Fact]
    public void ThisWorkbook_prefix_and_Sheets_alias_are_accepted()
    {
        var s = One("ThisWorkbook.Sheets(\"Config\").Cells(2, 3).Value = 1");
        Assert.Equal("Config", s.SheetName);
        Assert.Equal("C2", s.Address);
        Assert.Equal("write", s.Mode);
    }

    [Fact]
    public void Unknown_sheet_name_is_unknownSheet()
    {
        var s = One("v = Worksheets(\"Nope\").Range(\"A1\")");
        Assert.Null(s.SheetName);
        Assert.Equal("unknownSheet", s.UnresolvedReason);
    }

    [Fact]
    public void Numeric_index_resolves_by_position()
    {
        Assert.Equal("Config", One("v = Sheets(2).Range(\"B2\")").SheetName);
        Assert.Equal("unknownSheet", One("v = Sheets(9).Range(\"B2\")").UnresolvedReason);
    }

    [Fact]
    public void Variable_index_is_dynamicSheet()
    {
        Assert.Equal("dynamicSheet", One("v = Sheets(i).Range(\"B2\")").UnresolvedReason);
    }

    [Fact]
    public void Codename_qualifier_resolves()
    {
        var s = One("Blad2.Range(\"A1:B2\").ClearContents");
        Assert.Equal("Config", s.SheetName);
        Assert.Equal("A1:B2", s.Address);
        Assert.Equal("write", s.Mode);
    }

    [Fact]
    public void Unqualified_range_in_sheet_module_is_that_sheet()
    {
        var s = One("Range(\"A1\").Value = 5", module: "Blad1", kind: "documentModule");
        Assert.Equal("Data", s.SheetName);
        Assert.Equal("write", s.Mode);
    }

    [Fact]
    public void Me_in_sheet_module_is_that_sheet()
    {
        Assert.Equal("Data", One("v = Me.Cells(1, 1)", module: "Blad1", kind: "documentModule").SheetName);
    }

    [Fact]
    public void Unqualified_range_in_ThisWorkbook_or_standard_module_is_activeSheet()
    {
        Assert.Equal("activeSheet", One("v = Range(\"A1\")", module: "ThisWorkbook", kind: "documentModule").UnresolvedReason);
        Assert.Equal("activeSheet", One("v = Cells(1, 1)").UnresolvedReason);
        Assert.Equal("activeSheet", One("v = ActiveSheet.Range(\"A1\")").UnresolvedReason);
    }

    [Fact]
    public void With_block_qualifies_leading_dot_members()
    {
        var sites = Resolve("With Worksheets(\"Data\")\n    .Range(\"A1\").Value = 1\n    v = .Cells(2, 2)\nEnd With");
        Assert.Equal(2, sites.Count);
        Assert.All(sites, s => Assert.Equal("Data", s.SheetName));
        Assert.Equal("write", sites[0].Mode);
        Assert.Equal("read", sites[1].Mode);
    }

    [Fact]
    public void Nested_with_uses_innermost()
    {
        var sites = Resolve("With Worksheets(\"Data\")\n    With Worksheets(\"Config\")\n        v = .Range(\"A1\")\n    End With\n    w = .Range(\"B1\")\nEnd With");
        Assert.Equal("Config", sites[0].SheetName);
        Assert.Equal("Data", sites[1].SheetName);
    }

    [Fact]
    public void With_on_a_range_records_the_range_and_leading_dot_value_writes_it()
    {
        var sites = Resolve("With Worksheets(\"Data\").Range(\"A1:C3\")\n    .Value = 0\n    .Font.Bold = True\nEnd With");
        Assert.Equal(3, sites.Count);
        Assert.All(sites, s => { Assert.Equal("Data", s.SheetName); Assert.Equal("A1:C3", s.Address); });
        Assert.Equal("write", sites[1].Mode);
        Assert.Equal("write", sites[2].Mode);
    }

    [Fact]
    public void Alias_set_once_resolves()
    {
        var sites = Resolve("Set ws = Worksheets(\"Config\")\nws.Range(\"D4\").Value = 2");
        var s = Assert.Single(sites);
        Assert.Equal("Config", s.SheetName);
        Assert.Equal("write", s.Mode);
    }

    [Fact]
    public void Alias_to_codename_resolves()
    {
        Assert.Equal("My Sheet", One("Set ws = Blad3\nv = ws.Range(\"A1\")").SheetName);
    }

    [Fact]
    public void Alias_reassigned_is_unresolved()
    {
        var sites = Resolve("Set ws = Worksheets(\"Data\")\nSet ws = Worksheets(\"Config\")\nv = ws.Range(\"A1\")");
        var s = Assert.Single(sites);
        Assert.Null(s.SheetName);
        Assert.Equal("aliasReassigned", s.UnresolvedReason);
    }

    [Fact]
    public void Unknown_alias_is_dynamicSheet()
    {
        Assert.Equal("dynamicSheet", One("v = other.Range(\"A1\")").UnresolvedReason);
    }

    [Fact]
    public void Defined_name_resolves_via_refersTo()
    {
        var s = One("v = Range(\"Total\")");
        Assert.Equal("definedName", s.TargetKind);
        Assert.Equal("Total", s.DefinedNameRef);
        Assert.Equal("Config", s.SheetName);
        Assert.Equal("C2", s.Address);
    }

    [Fact]
    public void Bracket_shorthand_and_quoted_sheet_in_refersTo()
    {
        var s = One("[Prices].ClearContents");
        Assert.Equal("definedName", s.TargetKind);
        Assert.Equal("My Sheet", s.SheetName);
        Assert.Equal("A1:B9", s.Address);
        Assert.Equal("write", s.Mode);
    }

    [Fact]
    public void Named_constant_has_no_sheet()
    {
        var s = One("v = Range(\"Rate\")");
        Assert.Equal("definedName", s.TargetKind);
        Assert.Null(s.SheetName);
    }

    [Fact]
    public void Unknown_literal_is_unknownName()
    {
        var s = One("v = Worksheets(\"Data\").Range(\"Bogus\")");
        Assert.Equal("unknownName", s.UnresolvedReason);
    }

    [Fact]
    public void Sheet_qualified_literal_resolves_the_sheet()
    {
        var s = One("v = Range(\"Config!B7\")");
        Assert.Equal("Config", s.SheetName);
        Assert.Equal("B7", s.Address);
    }

    [Fact]
    public void Cells_with_variables_is_dynamicCells()
    {
        var s = One("Worksheets(\"Data\").Cells(r, c).Value = x");
        Assert.Equal("dynamicCells", s.TargetKind);
        Assert.Null(s.Address);
        Assert.Equal("write", s.Mode);
    }

    [Fact]
    public void Range_of_cells_is_one_dynamic_site()
    {
        var s = One("v = Worksheets(\"Data\").Range(Cells(1, 1), Cells(r, c)).Value");
        Assert.Equal("dynamicCells", s.TargetKind);
        Assert.Equal("Data", s.SheetName);
    }

    [Fact]
    public void Columns_and_Rows_targets()
    {
        var sites = Resolve("Worksheets(\"Data\").Columns(\"A:B\").Delete\nv = Blad1.Rows(5).Value\nBlad1.Columns(3).AutoFit");
        Assert.Equal(("column", "A:B", "write"), (sites[0].TargetKind, sites[0].Address, sites[0].Mode));
        Assert.Equal(("row", "5", "read"), (sites[1].TargetKind, sites[1].Address, sites[1].Mode));
        Assert.Equal(("column", "C", "read"), (sites[2].TargetKind, sites[2].Address, sites[2].Mode));
    }

    [Fact]
    public void UsedRange_is_wholeSheet()
    {
        var s = One("Set rng = Worksheets(\"Data\").UsedRange");
        Assert.Equal("wholeSheet", s.TargetKind);
        Assert.Equal("Data", s.SheetName);
    }

    [Fact]
    public void Nested_range_members_count_once_for_the_outer_site()
    {
        var sites = Resolve("v = Worksheets(\"Data\").Range(\"A1:B2\").Cells(1, 1).Value");
        var s = Assert.Single(sites);
        Assert.Equal("A1:B2", s.Address);
    }

    [Fact]
    public void Copy_line_yields_read_source_and_write_destination()
    {
        var sites = Resolve("Worksheets(\"Data\").Range(\"A1\").Copy Worksheets(\"Config\").Range(\"B1\")");
        Assert.Equal(2, sites.Count);
        Assert.Equal(("Data", "read"), (sites[0].SheetName, sites[0].Mode));
        Assert.Equal(("Config", "write"), (sites[1].SheetName, sites[1].Mode));
    }

    [Fact]
    public void If_comparison_is_not_a_write()
    {
        Assert.Equal("read", One("If Worksheets(\"Data\").Range(\"A1\").Value = 1 Then x = 2").Mode);
    }

    [Fact]
    public void Named_argument_and_comparison_operators_are_not_assignments()
    {
        Assert.Equal("read", One("Call Foo(rng:=Worksheets(\"Data\").Range(\"A1\"))").Mode);
        Assert.Equal("read", One("ok = Worksheets(\"Data\").Range(\"A1\").Value >= 3").Mode);
    }

    [Fact]
    public void Comment_and_string_contents_are_ignored()
    {
        Assert.Empty(Resolve("' Worksheets(\"Data\").Range(\"A1\")\nMsgBox \"see Range(\"\"A1\"\")\""));
    }

    [Fact]
    public void Line_number_and_procedure_are_reported()
    {
        var s = One("\nv = Blad1.Range(\"A1\")");
        Assert.Equal("Module1", s.Module);
        Assert.Equal("P", s.Procedure);
        Assert.Equal(3, s.Line);
    }
}
