using ModelContextProtocol;

namespace McpOffice.Tests;

public class ToolErrorTests
{
    [Fact]
    public void FileNotFound_returns_McpException_with_code_in_message()
    {
        var ex = Assert.IsType<McpException>(ToolError.FileNotFound(@"C:\missing.docx"));

        Assert.Contains("[file_not_found]", ex.Message);
        Assert.Contains(@"C:\missing.docx", ex.Message);
    }
}
