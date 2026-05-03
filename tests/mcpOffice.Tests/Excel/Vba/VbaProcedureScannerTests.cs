using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaProcedureScannerTests
{
    private static IReadOnlyList<ScannedProcedure> Scan(string moduleKind, string moduleName, string source)
    {
        var lines = VbaLineCleaner.Clean(source);
        return VbaProcedureScanner.Scan(moduleKind, moduleName, lines);
    }

    [Fact]
    public void Detects_simple_sub()
    {
        var procs = Scan("standardModule", "Module1",
            "Public Sub DoIt()\nEnd Sub");
        Assert.Single(procs);
        Assert.Equal("DoIt", procs[0].Procedure.Name);
        Assert.Equal("Module1.DoIt", procs[0].Procedure.FullyQualifiedName);
        Assert.Equal("Sub", procs[0].Procedure.Kind);
        Assert.Equal("Public", procs[0].Procedure.Scope);
    }

    [Fact]
    public void Detects_function_with_return_type()
    {
        var procs = Scan("standardModule", "Module1",
            "Function Add(a As Long, b As Long) As Long\nAdd = a + b\nEnd Function");
        var p = procs.Single().Procedure;
        Assert.Equal("Function", p.Kind);
        Assert.Equal("Long", p.ReturnType);
        Assert.Equal(2, p.Parameters.Count);
        Assert.Equal("a", p.Parameters[0].Name);
        Assert.Equal("Long", p.Parameters[0].Type);
    }

    [Fact]
    public void Detects_property_get()
    {
        var procs = Scan("classModule", "MyClass",
            "Public Property Get Name() As String\nEnd Property");
        Assert.Equal("PropertyGet", procs.Single().Procedure.Kind);
    }

    [Fact]
    public void Parses_optional_byval_with_default()
    {
        var procs = Scan("standardModule", "M",
            "Sub F(Optional ByVal x As String = \"d\")\nEnd Sub");
        var p = procs.Single().Procedure.Parameters.Single();
        Assert.True(p.Optional);
        Assert.False(p.ByRef);
        Assert.Equal("x", p.Name);
        Assert.Equal("String", p.Type);
        Assert.NotNull(p.DefaultValue);
    }

    [Fact]
    public void Detects_event_handler_in_document_module()
    {
        var procs = Scan("documentModule", "ThisWorkbook",
            "Private Sub Workbook_Open()\nEnd Sub");
        var p = procs.Single().Procedure;
        Assert.True(p.IsEventHandler);
        Assert.Equal("Workbook", p.EventTarget);
    }

    [Fact]
    public void Standard_module_with_underscore_name_is_not_event_handler()
    {
        var procs = Scan("standardModule", "Utils",
            "Sub Foo_Bar()\nEnd Sub");
        Assert.False(procs.Single().Procedure.IsEventHandler);
    }

    [Fact]
    public void Records_line_range()
    {
        var procs = Scan("standardModule", "M",
            "Sub A()\nx = 1\nEnd Sub");
        var p = procs.Single().Procedure;
        Assert.Equal(1, p.LineStart);
        Assert.Equal(3, p.LineEnd);
    }

    [Fact]
    public void Defaults_scope_to_null_when_unspecified()
    {
        var procs = Scan("standardModule", "M", "Sub A()\nEnd Sub");
        Assert.Null(procs.Single().Procedure.Scope);
    }

    [Fact]
    public void Multiple_procedures()
    {
        var procs = Scan("standardModule", "M",
            "Sub A()\nEnd Sub\n\nSub B()\nEnd Sub");
        Assert.Equal(2, procs.Count);
        Assert.Equal("A", procs[0].Procedure.Name);
        Assert.Equal("B", procs[1].Procedure.Name);
    }
}
