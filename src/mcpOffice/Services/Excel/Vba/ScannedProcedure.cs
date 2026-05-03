using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal sealed record ScannedProcedure(
    ExcelVbaProcedure Procedure,
    int CleanedLineStartIndex,    // index into the CleanedLine list (inclusive, line after the Sub/Function header)
    int CleanedLineEndIndex);     // index into the CleanedLine list (inclusive, line before End Sub/Function). May be less than CleanedLineStartIndex for empty-body procedures (e.g. "Sub A()\nEnd Sub"). Callers must guard endIdx >= startIdx before iterating.
