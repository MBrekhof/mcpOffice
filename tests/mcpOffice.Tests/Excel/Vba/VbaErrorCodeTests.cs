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
}
