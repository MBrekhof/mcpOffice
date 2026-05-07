namespace McpOffice.Models;

public sealed record ConversionHintsSummary(
    int TotalProcedures,
    int HintedProcedures,
    int ModuleCount,
    string? TargetParadigm,
    long WallTimeMs);
