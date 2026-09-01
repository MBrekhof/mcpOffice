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

    // VBA-007: real-world workbooks use mdl/bas as well (Air.xlsm: mdlAIR, mdlBalans).
    [Theory]
    [InlineData("mdlAIR", "AIR")]
    [InlineData("mdlBalans", "Balans")]
    [InlineData("basUtils", "Utils")]
    [InlineData("modules", "Modules")]   // lower-case continuation: not a prefix, PascalCased as a whole
    public void Naming_strips_mdl_and_bas_prefixes(string module, string expectedClass)
    {
        var s = ParadigmOverlayApplier.Apply(
            module: module, procedureName: "Run",
            scope: "Public", axes: Axes(), paradigm: "classLibrary");
        Assert.Equal(expectedClass, s.SuggestedClassName);
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

    [Fact]
    public void ClassLib_pure_leaf_static_method_no_blockers()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "P", "Public",
            Axes(purity: "pure", shape: "leaf"), "classLibrary");
        Assert.Equal("staticMethod", s.TargetType);
        Assert.Equal("static", s.Lifetime);
        Assert.Empty(s.Blockers);
    }

    [Fact]
    public void ClassLib_readsState_no_deps_static_method()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "P", "Public",
            Axes(purity: "readsState", shape: "leaf"), "classLibrary");
        Assert.Equal("staticMethod", s.TargetType);
    }

    [Fact]
    public void ClassLib_sideEffectful_database_instance_method_with_blocker()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "P", "Public",
            Axes(purity: "sideEffectful", shape: null, dependencies: new[] { "database" }),
            "classLibrary");
        Assert.Equal("instanceMethod", s.TargetType);
        Assert.Equal("scoped", s.Lifetime);
        Assert.Contains("requires_external_dependency_injection", s.Blockers);
    }

    [Fact]
    public void ClassLib_eventHandler_requires_manual_review()
    {
        var s = ParadigmOverlayApplier.Apply("Sheet1", "Worksheet_Change", "Public",
            Axes(trigger: "eventHandler"), "classLibrary");
        Assert.Equal("requiresManualReview", s.TargetType);
        Assert.Null(s.Lifetime);
        Assert.Contains("event_handler_no_pure_classlib_target", s.Blockers);
    }

    [Fact]
    public void Worker_macroEntryPoint_writesState_becomes_backgroundService()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "Main", "Public",
            Axes(trigger: "macroEntryPoint", purity: "writesState"), "workerService");
        Assert.Equal("backgroundService", s.TargetType);
        Assert.Equal("singleton", s.Lifetime);
    }

    [Fact]
    public void Worker_eventHandler_workbook_open_becomes_backgroundService()
    {
        var s = ParadigmOverlayApplier.Apply("ThisWorkbook", "Workbook_Open", "Public",
            Axes(trigger: "eventHandler"), "workerService");
        Assert.Equal("backgroundService", s.TargetType);
    }

    [Fact]
    public void Worker_eventHandler_auto_open_becomes_backgroundService()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "Auto_Open", "Public",
            Axes(trigger: "eventHandler"), "workerService");
        Assert.Equal("backgroundService", s.TargetType);
    }

    [Fact]
    public void Worker_other_procedure_becomes_instanceMethod()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "Helper", "Public",
            Axes(trigger: "calledOnly", purity: "pure"), "workerService");
        Assert.Equal("instanceMethod", s.TargetType);
        Assert.Equal("scoped", s.Lifetime);
    }

    [Fact]
    public void Worker_appends_excel_object_model_blocker()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "Main", "Public",
            Axes(trigger: "macroEntryPoint", purity: "sideEffectful",
                 dependencies: new[] { "excelObjectModel" }),
            "workerService");
        Assert.Contains("depends_on_excel_object_model", s.Blockers);
    }

    [Fact]
    public void WebApi_macroEntryPoint_public_becomes_apiAction()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "GetOrders", "Public",
            Axes(trigger: "macroEntryPoint"), "webApi");
        Assert.Equal("apiAction", s.TargetType);
        Assert.Equal("scoped", s.Lifetime);
    }

    [Fact]
    public void WebApi_macroEntryPoint_private_becomes_instanceMethod()
    {
        // Should never happen (private cannot be macroEntryPoint per the trigger rule),
        // but the table still says "any other" for non-public-entry rows.
        var s = ParadigmOverlayApplier.Apply("Module1", "GetOrders", "Private",
            Axes(trigger: "calledOnly"), "webApi");
        Assert.Equal("instanceMethod", s.TargetType);
    }

    [Fact]
    public void WebApi_appends_excel_blocker()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "GetOrders", "Public",
            Axes(trigger: "macroEntryPoint", dependencies: new[] { "excelObjectModel" }),
            "webApi");
        Assert.Contains("depends_on_excel_object_model", s.Blockers);
    }

    [Fact]
    public void Console_macroEntryPoint_becomes_consoleEntryPoint()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "Main", "Public",
            Axes(trigger: "macroEntryPoint"), "console");
        Assert.Equal("consoleEntryPoint", s.TargetType);
        Assert.Null(s.Lifetime);
    }

    [Fact]
    public void Console_pure_helper_becomes_staticMethod()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "Helper", "Public",
            Axes(trigger: "calledOnly", purity: "pure"), "console");
        Assert.Equal("staticMethod", s.TargetType);
        Assert.Equal("static", s.Lifetime);
    }

    [Fact]
    public void Console_sideEffectful_helper_becomes_instanceMethod()
    {
        var s = ParadigmOverlayApplier.Apply("Module1", "Doer", "Public",
            Axes(trigger: "calledOnly", purity: "sideEffectful",
                 dependencies: new[] { "filesystem" }),
            "console");
        Assert.Equal("instanceMethod", s.TargetType);
        Assert.Equal("scoped", s.Lifetime);
    }
}
