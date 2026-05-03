namespace McpOffice;

public static class ErrorCode
{
    public const string FileNotFound = "file_not_found";
    public const string FileExists = "file_exists";
    public const string InvalidPath = "invalid_path";
    public const string UnsupportedFormat = "unsupported_format";
    public const string ParseError = "parse_error";
    public const string IndexOutOfRange = "index_out_of_range";
    public const string SheetNotFound = "sheet_not_found";
    public const string RangeTooLarge = "range_too_large";
    public const string MergeFieldMissing = "merge_field_missing";
    public const string IoError = "io_error";
    public const string InternalError = "internal_error";
    public const string VbaProjectMissing = "vba_project_missing";
    public const string VbaProjectLocked = "vba_project_locked";
    public const string VbaParseError = "vba_parse_error";
    public const string ModuleNotFound = "module_not_found";
}
