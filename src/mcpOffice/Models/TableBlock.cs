namespace McpOffice.Models;

public sealed record TableBlock(int Index, IReadOnlyList<IReadOnlyList<string>> Rows);
