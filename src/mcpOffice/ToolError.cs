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
        Throw(ErrorCode.UnsupportedFormat, $"Unsupported format: {format}. Use one of pdf, html, rtf, txt, markdown, docx, odt.");

    public static Exception ParseError(string path, string detail) =>
        Throw(ErrorCode.ParseError, $"Could not parse {path}: {detail}");

    public static Exception IndexOutOfRange(int index, int max) =>
        Throw(ErrorCode.IndexOutOfRange, $"Index {index} is out of range (0..{max}).");

    public static Exception SheetNotFound(string sheet) =>
        Throw(ErrorCode.SheetNotFound, $"Worksheet not found: {sheet}");

    public static Exception RangeTooLarge(string range, int cellCount, int maxCells) =>
        Throw(ErrorCode.RangeTooLarge, $"Range {range} contains {cellCount} cells, which exceeds maxCells={maxCells}.");

    public static Exception RangeTooLargeRows(string range, int rowCount, int maxRows) =>
        Throw(ErrorCode.RangeTooLarge, $"Range {range} contains {rowCount} rows, which exceeds maxRows={maxRows}.");

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

    public static Exception ProcedureNotFound(string procedureName, IEnumerable<string> candidates) =>
        Throw(ErrorCode.ProcedureNotFound, $"Procedure not found: {procedureName}. Candidates: {string.Join(", ", candidates)}");

    public static Exception GraphTooLarge(int actualNodes, int maxNodes, string suggestion) =>
        Throw(ErrorCode.GraphTooLarge, $"Filtered graph has {actualNodes} nodes (maxNodes={maxNodes}). {suggestion}");

    public static Exception InvalidRenderOption(string detail) =>
        Throw(ErrorCode.InvalidRenderOption, $"Invalid render option: {detail}");

    public static Exception UnsupportedParadigm(string paradigm, IEnumerable<string> supported) =>
        Throw(ErrorCode.UnsupportedParadigm,
            $"Unsupported targetParadigm: {paradigm}. Supported values: {string.Join(", ", supported)}");

    public static Exception PasswordRequired(string path) =>
        Throw(ErrorCode.PasswordRequired, $"PDF is encrypted and needs a password: {path}");

    public static Exception PageNotFound(int pageNumber, int pageCount) =>
        Throw(ErrorCode.PageNotFound, $"Page {pageNumber} does not exist (document has {pageCount} page(s), numbered from 1).");

    public static Exception InvalidPageRange(string range, string detail) =>
        Throw(ErrorCode.InvalidPageRange, $"Could not parse pageRange '{range}': {detail}");

    public static Exception UnsupportedImageFormat(string format) =>
        Throw(ErrorCode.UnsupportedFormat, $"Unsupported image format: {format}. Use one of png, jpeg, jpg, bmp, gif, tiff.");

    private static McpException Throw(string code, string message) =>
        new($"[{code}] {message}");
}
