namespace McpOffice.Models;

public sealed record ConversionHints(
    ConversionHintsSummary Summary,
    IReadOnlyList<ProcedureHint> ProcedureHints,
    IReadOnlyList<ModuleCoupling> ModuleCoupling,
    IReadOnlyList<CouplingPair> CouplingPairs);
