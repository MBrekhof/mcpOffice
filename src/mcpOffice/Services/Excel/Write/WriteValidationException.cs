namespace McpOffice.Services.Excel.Write;

/// <summary>
/// Bad caller input found while building a workbook. The service layer translates it to
/// <c>[invalid_operation] {where}: {message}</c>, adding the operation or sheet index it knows.
/// </summary>
internal sealed class WriteValidationException(string message) : Exception(message);
