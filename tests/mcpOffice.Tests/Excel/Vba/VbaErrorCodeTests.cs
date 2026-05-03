using FluentAssertions;
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
        Action act = () => throw ToolError.ProcedureNotFound("ReadExports", new[] { "SaveDB", "Paste2Cell" });
        act.Should().Throw<ModelContextProtocol.McpException>()
           .Which.Message.Should().Contain("procedure_not_found")
           .And.Contain("ReadExports")
           .And.Contain("SaveDB")
           .And.Contain("Paste2Cell");
    }

    [Fact]
    public void GraphTooLarge_throws_McpException_with_count_and_max()
    {
        Action act = () => throw ToolError.GraphTooLarge(425, 300);
        act.Should().Throw<ModelContextProtocol.McpException>()
           .Which.Message.Should().Contain("graph_too_large")
           .And.Contain("425")
           .And.Contain("300");
    }

    [Fact]
    public void InvalidRenderOption_throws_McpException_with_option_and_message()
    {
        Action act = () => throw ToolError.InvalidRenderOption("format", "svg", "Use one of mermaid, dot.");
        act.Should().Throw<ModelContextProtocol.McpException>()
           .Which.Message.Should().Contain("invalid_render_option")
           .And.Contain("format")
           .And.Contain("svg")
           .And.Contain("mermaid");
    }
}
