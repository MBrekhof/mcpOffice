using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaReferenceCollectorTests
{
    private static (IReadOnlyList<ExcelVbaObjectModelRef> Om, IReadOnlyList<ExcelVbaDependency> Deps) Collect(
        string moduleName, string moduleKind, string source)
    {
        var lines = VbaLineCleaner.Clean(source);
        var procs = VbaProcedureScanner.Scan(moduleKind, moduleName, lines);
        var om = new List<ExcelVbaObjectModelRef>();
        var deps = new List<ExcelVbaDependency>();
        VbaReferenceCollector.Collect(moduleName, lines, procs, om, deps);
        return (om, deps);
    }

    [Fact]
    public void Captures_worksheets_with_literal()
    {
        var (om, _) = Collect("M", "standardModule",
            "Sub A()\nSet ws = Worksheets(\"Data\")\nEnd Sub");
        var r = Assert.Single(om);
        Assert.Equal("Worksheets", r.Api);
        Assert.Equal("Data", r.Literal);
    }

    [Fact]
    public void Captures_range()
    {
        var (om, _) = Collect("M", "standardModule",
            "Sub A()\nRange(\"A1:B10\").Value = 0\nEnd Sub");
        Assert.Equal("Range", om.Single().Api);
        Assert.Equal("A1:B10", om.Single().Literal);
    }

    [Fact]
    public void File_open_classified_as_file()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nOpen \"C:\\f.txt\" For Input As #1\nEnd Sub");
        var d = Assert.Single(deps);
        Assert.Equal("file", d.Kind);
    }

    [Fact]
    public void Adodb_classified_as_database()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nSet c = CreateObject(\"ADODB.Connection\")\nEnd Sub");
        var d = Assert.Single(deps);
        Assert.Equal("database", d.Kind);
        Assert.Equal("ADODB.Connection", d.Target);
    }

    [Fact]
    public void Msxml_classified_as_network()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nSet h = CreateObject(\"MSXML2.XMLHTTP\")\nEnd Sub");
        Assert.Equal("network", deps.Single().Kind);
    }

    [Fact]
    public void Outlook_falls_back_to_automation()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nSet o = CreateObject(\"Outlook.Application\")\nEnd Sub");
        var d = Assert.Single(deps);
        Assert.Equal("automation", d.Kind);
        Assert.Equal("Outlook.Application", d.Target);
    }

    [Fact]
    public void Shell_classified_as_shell()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nShell(\"notepad.exe\")\nEnd Sub");
        Assert.Equal("shell", deps.Single().Kind);
    }
}
