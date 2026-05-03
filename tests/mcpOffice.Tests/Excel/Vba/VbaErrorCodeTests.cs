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
    public void ProcedureNotFound_throws_McpException_with_code_and_candidates()
    {
        var ex = Assert.Throws<McpException>((Action)(() => throw ToolError.ProcedureNotFound("ReadExports", new[] { "SaveDB", "Paste2Cell" })));
        Assert.Contains("procedure_not_found", ex.Message);
        Assert.Contains("ReadExports", ex.Message);
        Assert.Contains("SaveDB", ex.Message);
        Assert.Contains("Paste2Cell", ex.Message);
    }

    [Fact]
    public void GraphTooLarge_throws_McpException_with_count_and_max()
    {
        var ex = Assert.Throws<McpException>((Action)(() => throw ToolError.GraphTooLarge(425, 300)));
        Assert.Contains("graph_too_large", ex.Message);
        Assert.Contains("425", ex.Message);
        Assert.Contains("300", ex.Message);
    }

    [Fact]
    public void InvalidRenderOption_throws_McpException_with_option_and_message()
    {
        var ex = Assert.Throws<McpException>((Action)(() => throw ToolError.InvalidRenderOption("format", "svg", "Use one of mermaid, dot.")));
        Assert.Contains("invalid_render_option", ex.Message);
        Assert.Contains("format", ex.Message);
        Assert.Contains("svg", ex.Message);
        Assert.Contains("mermaid", ex.Message);
    }
}
