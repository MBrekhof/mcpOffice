namespace McpOffice.Models;

public sealed record ExcelVbaAnalysisSummary(
    int ModuleCount,
    int ParsedModuleCount,
    int UnparsedModuleCount,
    int ProcedureCount,
    int EventHandlerCount,
    int CallEdgeCount,
    int ObjectModelReferenceCount,
    int DependencyCount);
