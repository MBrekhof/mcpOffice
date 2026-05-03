namespace McpOffice.Models;

public sealed record ExcelVbaAnalysis(
    bool HasVbaProject,
    ExcelVbaAnalysisSummary Summary,
    IReadOnlyList<ExcelVbaModuleAnalysis>? Modules,
    IReadOnlyList<ExcelVbaCallEdge>? CallGraph,
    ExcelVbaReferences? References);
