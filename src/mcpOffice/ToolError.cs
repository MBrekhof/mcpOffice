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

    public static Exception SheetNotFound(string sheet) =>
        Throw(ErrorCode.SheetNotFound, $"Worksheet not found: {sheet}");

    public static Exception RangeTooLarge(string range, int cellCount, int maxCells) =>
        Throw(ErrorCode.RangeTooLarge, $"Range {range} contains {cellCount} cells, which exceeds maxCells={maxCells}.");

    public static Exception MergeFieldMissing(IEnumerable<string> fields) =>
        Throw(ErrorCode.MergeFieldMissing, $"Template fields with no value in dataJson: {string.Join(", ", fields)}");

    public static Exception IoError(string detail) =>
        Throw(ErrorCode.IoError, $"IO error: {detail}");

    public static Exception Internal(string detail) =>
        Throw(ErrorCode.InternalError, $"Internal error: {detail}");

    public static Exception VbaProjectMissing(string path) =>
        Throw(ErrorCode.VbaProjectMissing, $"No VBA project in workbook: {path}");

    public static Exception VbaProjectLocked(string path) =>
        Throw(ErrorCode.VbaProjectLocked, $"VBA project is locked for viewing: {path}");

    public static Exception VbaParseError(string path, string detail) =>
        Throw(ErrorCode.VbaParseError, $"Could not parse VBA project in {path}: {detail}");

    public static Exception ModuleNotFound(string moduleName, IEnumerable<string> available) =>
        Throw(ErrorCode.ModuleNotFound, $"Module not found: {moduleName}. Available modules: {string.Join(", ", available)}");

    public static Exception ProcedureNotFound(string procedureName, IEnumerable<string> available) =>
        Throw(ErrorCode.ProcedureNotFound, $"Procedure not found: {procedureName}. Available procedures: {string.Join(", ", available)}");

    public static Exception GraphTooLarge(int actualNodeCount, int maxNodes) =>
        Throw(ErrorCode.GraphTooLarge, $"Filtered call graph has {actualNodeCount} nodes, which exceeds maxNodes={maxNodes}. Narrow the result with moduleName, procedureName, or a smaller depth.");

    public static Exception InvalidRenderOption(string optionName, string value, string detail) =>
        Throw(ErrorCode.InvalidRenderOption, $"Invalid value for {optionName}: '{value}'. {detail}");

    private static McpException Throw(string code, string message) =>
        new($"[{code}] {message}");
}
