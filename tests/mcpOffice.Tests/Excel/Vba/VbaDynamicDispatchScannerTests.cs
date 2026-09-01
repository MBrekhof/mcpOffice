using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaDynamicDispatchScannerTests
{
    private static IReadOnlyList<VbaDynamicDispatchScanner.DynamicDispatch> Scan(string statement)
    {
        var lines = VbaLineCleaner.Clean($"Sub A()\n{statement}\nEnd Sub");
        var procs = VbaProcedureScanner.Scan("standardModule", "M", lines);
        return VbaDynamicDispatchScanner.Scan("M", lines, procs);
    }

    [Theory]
    [InlineData("Application.OnTime Now + TimeValue(\"00:00:10\"), \"Tick\"", "OnTime", "Tick")]
    [InlineData("Application.OnTime Now, \"Module1.Tick\"", "OnTime", "Module1.Tick")]
    [InlineData("Application.OnTime EarliestTime:=Now, Procedure:=\"Tick\"", "OnTime", "Tick")]
    [InlineData("Application.OnTime Now, \"Tick\", , False", "OnTime", "Tick")]
    [InlineData("Application.OnTime Now, procName", "OnTime", null)]
    [InlineData("Application.OnKey \"^q\", \"Quit\"", "OnKey", "Quit")]
    [InlineData("Application.OnKey Key:=\"^q\", Procedure:=\"Quit\"", "OnKey", "Quit")]
    [InlineData("Application.OnKey \"^q\", target", "OnKey", null)]
    [InlineData("Application.Run \"Proc\"", "Run", "Proc")]
    [InlineData("Application.Run(\"Proc\", 1, 2)", "Run", "Proc")]
    [InlineData("x = Application.Run(\"'Book.xlsm'!Proc\")", "Run", "'Book.xlsm'!Proc")]
    [InlineData("Application.Run Macro:=\"Proc\"", "Run", "Proc")]
    [InlineData("Run \"Proc\"", "Run", "Proc")]
    [InlineData("Call Run(\"Proc\")", "Run", "Proc")]
    [InlineData("Application.Run macroName", "Run", null)]
    [InlineData("Application.Run \"Proc\" ' trailing comment", "Run", "Proc")]
    [InlineData("btn.OnAction = \"Module1.Clicked\"", "OnAction", "Module1.Clicked")]
    [InlineData(".OnAction = \"Clicked\"", "OnAction", "Clicked")]
    [InlineData(".OnAction = sName", "OnAction", null)]
    [InlineData("CallByName obj, \"DoIt\", VbMethod", "CallByName", "DoIt")]
    [InlineData("r = CallByName(obj, \"Value\", VbGet)", "CallByName", "Value")]
    [InlineData("CallByName obj, propName, VbGet", "CallByName", null)]
    public void Detects_each_api_with_literal_or_null_target(string statement, string api, string? target)
    {
        var d = Assert.Single(Scan(statement));
        Assert.Equal(api, d.Api);
        Assert.Equal(target, d.TargetLiteral);
    }

    [Theory]
    [InlineData("Application.OnKey \"^q\"")]                 // one-arg form resets the key
    [InlineData("Shell \"cmd\"")]
    [InlineData("' Application.OnTime Now, \"Tick\"")]      // comment
    [InlineData("MsgBox \"Application.Run\"")]              // inside a string literal
    [InlineData("RunQuery \"x\"")]
    [InlineData("Run = 5")]
    [InlineData("wsh.Run \"cmd.exe\"")]                     // WScript.Shell, not Application
    public void Ignores_non_dispatch_lines(string statement)
    {
        Assert.Empty(Scan(statement));
    }

    [Fact]
    public void Records_module_procedure_and_line()
    {
        var d = Assert.Single(Scan("Application.Run \"Proc\""));
        Assert.Equal("M", d.Module);
        Assert.Equal("A", d.Procedure);
        Assert.Equal(2, d.Line);
    }

    [Fact]
    public void Ignores_lines_outside_procedures()
    {
        var lines = VbaLineCleaner.Clean("Application.Run \"Proc\"\nSub A()\nEnd Sub");
        var procs = VbaProcedureScanner.Scan("standardModule", "M", lines);
        Assert.Empty(VbaDynamicDispatchScanner.Scan("M", lines, procs));
    }
}
