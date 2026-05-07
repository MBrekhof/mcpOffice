using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class ParadigmOverlayApplierTests
{
    private static ProcedureAxes Axes(
        string trigger = "calledOnly",
        string purity = "pure",
        string? shape = "leaf",
        params string[] dependencies) =>
        new(trigger, purity, shape, dependencies);

    [Fact]
    public void Naming_strips_mod_prefix()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "modOrders", procedureName: "ProcessOrder",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("Orders", s.SuggestedClassName);
    }

    [Fact]
    public void Naming_strips_cls_prefix()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "clsCustomer", procedureName: "GetById",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("Customer", s.SuggestedClassName);
    }

    [Fact]
    public void Naming_strips_frm_prefix()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "frmLogin", procedureName: "Validate",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("Login", s.SuggestedClassName);
    }

    [Fact]
    public void Naming_passes_through_when_no_prefix()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "Module1", procedureName: "Main",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("Module1", s.SuggestedClassName);
    }

    [Fact]
    public void Naming_method_pascal_case()
    {
        var s = ParadigmOverlayApplier.Apply(
            module: "Module1", procedureName: "do_thing",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal("DoThing", s.SuggestedMethodName);
    }

    [Fact]
    public void IsPublic_mirrors_scope()
    {
        var pub = ParadigmOverlayApplier.Apply(
            "M", "P", scope: "Public", axes: Axes(), paradigm: "classLibrary");
        var priv = ParadigmOverlayApplier.Apply(
            "M", "P", scope: "Private", axes: Axes(), paradigm: "classLibrary");
        Assert.True(pub.IsPublic);
        Assert.False(priv.IsPublic);
    }
}
