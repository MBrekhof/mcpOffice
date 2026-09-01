using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

/// <summary>Gated real-world checks; no-ops on machines without the sample corpus.</summary>
public class VbaEntryPointsSampleTests
{
    private const string Ring = @"C:\Projects\mcpOffice-samples\RingOnderzoek.xlsm";
    private const string Air  = @"C:\Projects\mcpOffice-samples\Air.xlsm";

    [Fact]
    public void Synthetic_fixture_events_are_entry_points_and_class_method_is_dead()
    {
        // Unconditional: synthetic-vba.xlsm ships in tests/fixtures (see its README).
        var r = new ExcelWorkbookService().ListVbaEntryPoints(
            TestFixtures.Path("synthetic-vba.xlsm"), includeUnreachable: true, moduleName: null);

        Assert.True(r.HasVbaProject);
        Assert.Contains(r.EntryPoints, e => e.Kind == "eventHandler" && e.Procedure == "ThisWorkbook.Workbook_Open");
        Assert.DoesNotContain(r.Unreachable!, u => u.Procedure == "Module1.Main");   // Workbook_Open calls Main
        var greet = Assert.Single(r.Unreachable!, u => u.Procedure == "Class1.Greet");
        Assert.Equal("medium", greet.Confidence);   // Public class member: could be reached through an object variable
    }

    [Fact]
    public void RingOnderzoek_pictures_wired_to_macros_resolve()
    {
        if (!File.Exists(Ring)) return;
        var r = new ExcelWorkbookService().ListVbaEntryPoints(Ring, includeUnreachable: true, moduleName: null);

        Assert.True(r.HasVbaProject);
        // The only drawing macro in this workbook is a connector wired to a macro with an argument.
        var copy = Assert.Single(r.EntryPoints, e => e.Kind == "shapeMacro");
        Assert.Equal("'Copy_results(2)'", copy.Target);
        Assert.True(r.Summary.ReachableCount > 0);
        Assert.Equal(r.Summary.ProcedureCount, r.Summary.ReachableCount + r.Summary.UnreachableCount);
    }

    [Fact]
    public void Air_form_control_buttons_resolve_and_dead_code_is_bounded()
    {
        if (!File.Exists(Air)) return;
        var r = new ExcelWorkbookService().ListVbaEntryPoints(Air, includeUnreachable: true, moduleName: null);

        Assert.True(r.HasVbaProject);
        Assert.Contains(r.EntryPoints, e => e.Kind == "shapeMacro" && e.Target == "[0]!GetILIS" && e.Resolved);
        Assert.Contains(r.EntryPoints, e => e.Kind == "shapeMacro" && e.Target == "[0]!Inlezen" && e.Resolved);
        Assert.Contains(r.EntryPoints, e => e.Kind == "formControlMacro" && e.Target == "[0]!NextDate" && e.Resolved);
        Assert.Equal(0, r.Summary.SkippedDrawingParts);
        Assert.True(r.Summary.UnreachableCount < r.Summary.ProcedureCount, "not everything can be dead");
        Assert.All(r.Unreachable!, u => Assert.Contains(u.Confidence, new[] { "high", "medium" }));
    }
}
