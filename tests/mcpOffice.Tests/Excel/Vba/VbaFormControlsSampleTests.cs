using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

/// <summary>Gated real-world checks for excel_list_vba_form_controls.</summary>
public class VbaFormControlsSampleTests
{
    private const string OlieGc = @"C:\Projects\mcpOffice-samples\OlieGC - LABWARE PRD.xlsm";
    private const string Qqq2   = @"C:\Projects\mcpOffice-samples\QQQ2 - Absolute.xlsm";

    [Fact]
    public void Synthetic_fixture_without_forms_returns_empty_forms()
    {
        var r = new ExcelWorkbookService().ListVbaFormControls(TestFixtures.Path("synthetic-vba.xlsm"), null);
        Assert.True(r.HasVbaProject);
        Assert.Empty(r.Forms);
    }

    [Fact]
    public void OlieGC_forms_have_typed_controls()
    {
        if (!File.Exists(OlieGc)) return;
        var r = new ExcelWorkbookService().ListVbaFormControls(OlieGc, null);
        Assert.True(r.HasVbaProject);
        Assert.NotEmpty(r.Forms);
        Assert.True(r.Summary.ControlCount > 0, "expected at least one inferred control");
        Assert.True(r.Summary.TypedControlCount > 0, "expected at least one typed control");
    }

    [Fact]
    public void QQQ2_password_form_is_listed()
    {
        if (!File.Exists(Qqq2)) return;
        var r = new ExcelWorkbookService().ListVbaFormControls(Qqq2, "frmPwd");
        var form = Assert.Single(r.Forms);
        Assert.Equal("frmPwd", form.Name);
    }
}
