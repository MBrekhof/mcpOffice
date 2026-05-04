using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class VbaErrorCodeTests
{
    [Fact]
    public void VbaProjectMissing_has_stable_code_and_path()
    {
        var ex = ToolError.VbaProjectMissing(@"C:\book.xlsx");
        Assert.IsType<McpException>(ex);
        Assert.Contains("vba_project_missing", ex.Message);
        Assert.Contains(@"C:\book.xlsx", ex.Message);
    }

    [Fact]
    public void VbaProjectLocked_has_stable_code_and_path()
    {
        var ex = ToolError.VbaProjectLocked(@"C:\book.xlsm");
        Assert.IsType<McpException>(ex);
        Assert.Contains("vba_project_locked", ex.Message);
        Assert.Contains(@"C:\book.xlsm", ex.Message);
    }

    [Fact]
    public void VbaParseError_has_stable_code_and_detail()
    {
        var ex = ToolError.VbaParseError(@"C:\book.xlsm", "bad chunk header");
        Assert.IsType<McpException>(ex);
        Assert.Contains("vba_parse_error", ex.Message);
        Assert.Contains("bad chunk header", ex.Message);
    }

    [Fact]
    public void ProcedureNotFound_has_stable_code_and_candidates()
    {
        var ex = ToolError.ProcedureNotFound("ReadExports", new[] { "SaveDB", "Paste2Cell" });
        Assert.IsType<McpException>(ex);
        Assert.Contains("procedure_not_found", ex.Message);
        Assert.Contains("ReadExports", ex.Message);
        Assert.Contains("SaveDB", ex.Message);
        Assert.Contains("Paste2Cell", ex.Message);
    }

    [Fact]
    public void GraphTooLarge_has_stable_code_and_counts()
    {
        var ex = ToolError.GraphTooLarge(523, 300, "Add moduleName to narrow the view.");
        Assert.IsType<McpException>(ex);
        Assert.Contains("graph_too_large", ex.Message);
        Assert.Contains("523", ex.Message);
        Assert.Contains("300", ex.Message);
        Assert.Contains("moduleName", ex.Message);
    }

    [Fact]
    public void InvalidRenderOption_has_stable_code_and_detail()
    {
        var ex = ToolError.InvalidRenderOption("procedureName requires moduleName");
        Assert.IsType<McpException>(ex);
        Assert.Contains("invalid_render_option", ex.Message);
        Assert.Contains("procedureName requires moduleName", ex.Message);
    }
}
