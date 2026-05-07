namespace McpOffice.Models;

public sealed record ProcedureAxes(
    string Trigger,
    string Purity,
    string? Shape,
    IReadOnlyList<string> Dependencies);
