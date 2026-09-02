using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaCorpusComparerTests
{
    // Subs, not Functions: a Function's return assignment carries its own name, so a renamed
    // Function is by construction a near-duplicate, not an identical body.
    private const string Utils = """
        Public Sub Clean(s As String)
            s = Trim(s)
            s = Replace(s, vbTab, " ")
            Debug.Print s
        End Sub
        Private Sub Tiny()
            x = 1
        End Sub
        """;

    private const string UtilsRenamed = """
        Public Sub Schoon(s As String)
            S = Trim(S)   ' comment must not matter
            s = Replace(s, vbTab, " ")
            Debug.Print s
        End Sub
        """;

    private const string UtilsPatched = """
        Public Function Clean(s As String) As String
            s = Trim(s)
            s = Replace(s, vbTab, " ")
            s = Replace(s, vbCr, "")
            s = Replace(s, vbLf, "")
            s = Replace(s, vbVerticalTab, "")
            s = Replace(s, "  ", " ")
            s = Replace(s, "   ", " ")
            s = Replace(s, "    ", " ")
            s = Replace(s, "     ", " ")
            Clean = s
        End Function
        """;

    private static VbaCorpusComparer.WorkbookInput Wb(string path, params (string Module, string Code)[] modules) =>
        new(path, new ExcelVbaProject(true, modules.Select(m => new ExcelVbaModule(m.Module, "standardModule", 1, m.Code)).ToList()), null);

    [Fact]
    public void Identical_bodies_group_across_workbooks_even_when_renamed()
    {
        var r = VbaCorpusComparer.Compare(
            [Wb(@"C:\a.xlsm", ("mdlUtils", Utils)), Wb(@"C:\b.xlsm", ("mdlUtils", Utils)), Wb(@"C:\c.xlsm", ("Helpers", UtilsRenamed))],
            minOccurrences: 2, maxProcedures: 200, includeNearDuplicates: true);

        var g = Assert.Single(r.SharedProcedures);
        Assert.Equal("identical", g.Tier);
        Assert.Equal("Clean", g.Name);   // most common name in the group
        Assert.Equal(3, g.Occurrences.Count);
        Assert.Contains(g.Occurrences, o => o.Workbook == @"C:\c.xlsm" && o.Procedure == "Schoon");
        Assert.Equal(1, r.Summary.IdenticalGroups);
        Assert.Equal(3, r.Summary.SharedProcedureCount);
    }

    [Fact]
    public void Tiny_bodies_are_ignored_as_noise()
    {
        var r = VbaCorpusComparer.Compare(
            [Wb(@"C:\a.xlsm", ("M", Utils)), Wb(@"C:\b.xlsm", ("M", Utils))], 2, 200, true);
        Assert.DoesNotContain(r.SharedProcedures, g => g.Name == "Tiny");
    }

    [Fact]
    public void Near_duplicates_group_by_name_when_similar_enough()
    {
        // 10-line body vs 4-line body: not ≥ 0.9 similar → no near-duplicate group.
        var far = VbaCorpusComparer.Compare(
            [Wb(@"C:\a.xlsm", ("M", Utils)), Wb(@"C:\b.xlsm", ("M", UtilsPatched))], 2, 200, true);
        Assert.Empty(far.SharedProcedures);

        // Same 10-line body with one line changed: 18/20 = 0.9 → near duplicate.
        // (Whitespace inside literals is collapsed like any other whitespace — change code, not spaces.)
        var patched2 = UtilsPatched.Replace("vbLf", "vbNewLine");
        var near = VbaCorpusComparer.Compare(
            [Wb(@"C:\a.xlsm", ("M", UtilsPatched)), Wb(@"C:\b.xlsm", ("M", patched2))], 2, 200, true);
        var g = Assert.Single(near.SharedProcedures);
        Assert.Equal("nearDuplicate", g.Tier);
        Assert.Equal(2, g.Occurrences.Count);
        Assert.Contains(g.Occurrences, o => o.Similarity < 1.0);
        Assert.Equal(1, near.Summary.NearDuplicateGroups);

        var off = VbaCorpusComparer.Compare(
            [Wb(@"C:\a.xlsm", ("M", UtilsPatched)), Wb(@"C:\b.xlsm", ("M", patched2))], 2, 200, includeNearDuplicates: false);
        Assert.Empty(off.SharedProcedures);
    }

    [Fact]
    public void Shared_module_is_reported_when_most_of_its_procedures_are_shared()
    {
        var r = VbaCorpusComparer.Compare(
            [Wb(@"C:\a.xlsm", ("mdlUtils", Utils)), Wb(@"C:\b.xlsm", ("mdlUtils", Utils))], 2, 200, true);
        var m = Assert.Single(r.SharedModules);
        Assert.Equal("mdlUtils", m.Module);
        Assert.Equal(2, m.Workbooks.Count);
        Assert.Equal(1.0, m.SharedProcedureRatio);
    }

    [Fact]
    public void MinOccurrences_and_maxProcedures_are_honoured()
    {
        var inputs = new[] { Wb(@"C:\a.xlsm", ("M", Utils)), Wb(@"C:\b.xlsm", ("M", Utils)) };
        Assert.Empty(VbaCorpusComparer.Compare(inputs, minOccurrences: 3, 200, true).SharedProcedures);

        var capped = VbaCorpusComparer.Compare(inputs, 2, maxProcedures: 0, true);
        Assert.Empty(capped.SharedProcedures);
        Assert.True(capped.Truncated);
    }

    [Fact]
    public void Failed_and_macro_free_workbooks_are_listed_not_fatal()
    {
        var r = VbaCorpusComparer.Compare(
        [
            Wb(@"C:\a.xlsm", ("M", Utils)),
            new VbaCorpusComparer.WorkbookInput(@"C:\locked.xlsm", null, "[vba_project_locked] locked"),
            new VbaCorpusComparer.WorkbookInput(@"C:\plain.xlsm", new ExcelVbaProject(false, []), null),
        ], 2, 200, true);

        Assert.Equal(3, r.Workbooks.Count);
        Assert.Equal("[vba_project_locked] locked", r.Workbooks[1].Error);
        Assert.False(r.Workbooks[2].HasVbaProject);
        Assert.Empty(r.SharedProcedures);
    }
}
