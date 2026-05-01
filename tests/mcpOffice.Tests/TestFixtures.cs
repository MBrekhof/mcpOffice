namespace McpOffice.Tests;

internal static class TestFixtures
{
    public static string Path(string name)
    {
        var asmDir = System.IO.Path.GetDirectoryName(typeof(TestFixtures).Assembly.Location)!;
        var dir = new DirectoryInfo(asmDir);
        while (dir is not null && !File.Exists(System.IO.Path.Combine(dir.FullName, "mcpOffice.sln")))
            dir = dir.Parent;
        if (dir is null) throw new InvalidOperationException("Could not locate repo root.");
        return System.IO.Path.Combine(dir.FullName, "tests", "fixtures", name);
    }
}
