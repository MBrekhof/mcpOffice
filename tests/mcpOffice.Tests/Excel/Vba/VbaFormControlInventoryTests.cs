using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaFormControlInventoryTests
{
    private const string Form = """
        Private WithEvents grid As MSForms.ListBox
        Dim rng As Range

        Private Sub UserForm_Initialize()
            cboType.AddItem "A"
            Me.txtName.Text = ""
            rng.Value = 1
        End Sub

        Private Sub cmdOK_Click()
            If Me.txtName.Text = "" Then Exit Sub
            lstItems.AddItem txtName.Text
            Me.Hide
        End Sub

        Private Sub Toggle_Click()
            Toggle.Caption = "x"
        End Sub

        Private Sub Picker_DropButtonClick()
        End Sub

        Private Sub Amount_Change()
            Amount.MaxLength = 5
        End Sub
        """;

    private static ExcelVbaUserForm Run(string code = Form) =>
        VbaFormControlInventory.Analyze(new ExcelVbaModule("frmOrder", "userForm", 1, code));

    private static ExcelVbaFormControl Control(string name, string code = Form) =>
        Assert.Single(Run(code).Controls, c => c.Name == name);

    [Fact]
    public void Handler_name_yields_control_with_event_and_prefix_type()
    {
        var c = Control("cmdOK");
        Assert.Equal(("CommandButton", "prefix"), (c.InferredType, c.TypeConfidence));
        Assert.Equal(["Click"], c.Events);
    }

    [Fact]
    public void Me_reference_collects_properties_and_prefix_type()
    {
        var c = Control("txtName");
        Assert.Equal("TextBox", c.InferredType);
        Assert.Equal(["Text"], c.ReferencedProperties);
    }

    [Fact]
    public void Bare_reference_counts_only_with_a_hungarian_prefix_or_known_name()
    {
        Assert.Equal("ComboBox", Control("cboType").InferredType);
        Assert.Equal("ListBox", Control("lstItems").InferredType);
        Assert.DoesNotContain(Run().Controls, c => c.Name == "rng");   // Dim rng As Range: a local, not a control
    }

    [Fact]
    public void Declared_msforms_type_wins()
    {
        var c = Control("grid");
        Assert.Equal(("ListBox", "declared"), (c.InferredType, c.TypeConfidence));
    }

    [Fact]
    public void Event_and_member_hints_type_unprefixed_controls()
    {
        Assert.Equal(("ComboBox", "event"), (Control("Picker").InferredType, Control("Picker").TypeConfidence));
        Assert.Equal(("TextBox", "member"), (Control("Amount").InferredType, Control("Amount").TypeConfidence));
        // Click + .Caption: Click says CommandButton, Caption says Label; member wins.
        Assert.Equal(("Label", "member"), (Control("Toggle").InferredType, Control("Toggle").TypeConfidence));
    }

    [Fact]
    public void Form_events_and_handler_count_are_reported_and_Me_Hide_is_not_a_control()
    {
        var f = Run();
        Assert.Equal(["Initialize"], f.FormEvents);
        Assert.Equal(5, f.HandlerCount);
        Assert.DoesNotContain(f.Controls, c => c.Name == "Hide");
        Assert.Equal("frmOrder", f.Name);
    }

    [Fact]
    public void Vbe_default_names_type_by_their_type_name_not_by_event()
    {
        // OlieGC's frmKeuze has Label2_Click: the Click hint said CommandButton, but the VBE named it a Label.
        const string code = "Private Sub Label2_Click()\nEnd Sub\nPrivate Sub ComboBox1_DropButtonClick()\nEnd Sub\nPrivate Sub TextBox1_Change()\nEnd Sub";
        Assert.Equal(("Label", "prefix"), (Control("Label2", code).InferredType, Control("Label2", code).TypeConfidence));
        Assert.Equal(("ComboBox", "prefix"), (Control("ComboBox1", code).InferredType, Control("ComboBox1", code).TypeConfidence));
        Assert.Equal(("TextBox", "prefix"), (Control("TextBox1", code).InferredType, Control("TextBox1", code).TypeConfidence));
    }

    [Fact]
    public void Controls_are_sorted_and_unknown_type_is_Control_none()
    {
        var f = Run("Private Sub Widget_Enter()\nEnd Sub\nPrivate Sub Alpha_Exit()\nEnd Sub");
        Assert.Equal(["Alpha", "Widget"], f.Controls.Select(c => c.Name));
        Assert.All(f.Controls, c => Assert.Equal(("Control", "none"), (c.InferredType, c.TypeConfidence)));
    }

    [Fact]
    public void Empty_form_has_no_controls()
    {
        var f = Run("");
        Assert.Empty(f.Controls);
        Assert.Equal(0, f.HandlerCount);
    }
}
