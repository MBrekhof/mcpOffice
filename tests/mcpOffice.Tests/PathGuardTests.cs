using ModelContextProtocol;

namespace McpOffice.Tests;

public class PathGuardTests
{
    [Fact]
    public void RequireAbsolute_rejects_relative_path()
    {
        var ex = Assert.Throws<McpException>(() => PathGuard.RequireAbsolute("foo.docx"));

        Assert.Contains("[invalid_path]", ex.Message);
    }

    [Fact]
    public void RequireAbsolute_accepts_absolute_path()
    {
        PathGuard.RequireAbsolute(@"C:\foo.docx");
    }

    [Fact]
    public void RequireExists_throws_when_file_is_missing()
    {
        var ex = Assert.Throws<McpException>(() => PathGuard.RequireExists(@"C:\definitely-does-not-exist-xyz.docx"));

        Assert.Contains("[file_not_found]", ex.Message);
    }

    [Fact]
    public void RequireWritable_throws_when_file_exists_without_overwrite()
    {
        var tmp = Path.Combine(Path.GetTempPath(), $"mcpoffice-pathguard-{Guid.NewGuid():N}.tmp");
        File.WriteAllText(tmp, "x");

        try
        {
            var ex = Assert.Throws<McpException>(() => PathGuard.RequireWritable(tmp, overwrite: false));
            Assert.Contains("[file_exists]", ex.Message);
        }
        finally
        {
            File.Delete(tmp);
        }
    }
}
