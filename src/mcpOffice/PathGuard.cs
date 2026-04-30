namespace McpOffice;

public static class PathGuard
{
    public static void RequireAbsolute(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path))
        {
            throw ToolError.InvalidPath(path);
        }
    }

    public static void RequireExists(string path)
    {
        RequireAbsolute(path);

        if (!File.Exists(path))
        {
            throw ToolError.FileNotFound(path);
        }
    }

    public static void RequireWritable(string path, bool overwrite)
    {
        RequireAbsolute(path);

        if (File.Exists(path) && !overwrite)
        {
            throw ToolError.FileExists(path);
        }

        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
        {
            Directory.CreateDirectory(dir);
        }
    }
}
