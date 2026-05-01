namespace McpOffice.Services.Excel.Vba;

internal sealed record VbaModuleEntry(
    string Name,
    string StreamName,
    uint TextOffset,
    ushort Type);
