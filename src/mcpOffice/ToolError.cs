using ModelContextProtocol;

namespace McpOffice;

public static class ToolError
{
    public static Exception FileNotFound(string path) =>
        Throw(ErrorCode.FileNotFound, $"File not found: {path}");

    public static Exception FileExists(string path) =>
        Throw(ErrorCode.FileExists, $"Output already exists (pass overwrite=true to replace): {path}");

    public static Exception InvalidPath(string path) =>
        Throw(ErrorCode.InvalidPath, $"Path must be absolute and well-formed: {path}");

    public static Exception UnsupportedFormat(string format) =>
        Throw(ErrorCode.UnsupportedFormat, $"Unsupported format: {format}. Use one of pdf, html, rtf, txt, markdown, docx.");

    public static Exception ParseError(string path, string detail) =>
        Throw(ErrorCode.ParseError, $"Could not parse {path}: {detail}");

    public static Exception IndexOutOfRange(int index, int max) =>
        Throw(ErrorCode.IndexOutOfRange, $"Index {index} is out of range (0..{max}).");

    public static Exception MergeFieldMissing(IEnumerable<string> fields) =>
        Throw(ErrorCode.MergeFieldMissing, $"Template fields with no value in dataJson: {string.Join(", ", fields)}");

    public static Exception IoError(string detail) =>
        Throw(ErrorCode.IoError, $"IO error: {detail}");

    public static Exception Internal(string detail) =>
        Throw(ErrorCode.InternalError, $"Internal error: {detail}");

    private static McpException Throw(string code, string message) =>
        new($"[{code}] {message}");
}
