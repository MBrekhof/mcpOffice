using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal sealed record ScannedProcedure(
    ExcelVbaProcedure Procedure,
    int CleanedLineStartIndex,    // index into the CleanedLine list (inclusive, line after the Sub/Function header)
    int CleanedLineEndIndex);     // index into the CleanedLine list (inclusive, line before End Sub/Function)
