namespace McpOffice.Models;

/// <summary>Result of <c>excel_list_vba_entry_points</c>: what actually runs in a workbook, and what never can.</summary>
public sealed record ExcelVbaEntryPointsResult(
    string Path,
    bool HasVbaProject,
    ExcelVbaEntryPointsSummary Summary,
    IReadOnlyList<ExcelVbaEntryPoint> EntryPoints,
    IReadOnlyList<ExcelVbaUnreachableProcedure>? Unreachable,
    bool Truncated);

public sealed record ExcelVbaEntryPointsSummary(
    int EntryPointCount,
    IReadOnlyDictionary<string, int> ByKind,
    int ProcedureCount,
    int ReachableCount,
    int UnreachableCount,
    int UnresolvedMacroReferences,
    int DynamicDispatchUnresolved,
    int SkippedDrawingParts);

/// <summary>
/// One way code gets started. <c>Kind</c> is a closed set: eventHandler, autoMacro, shapeMacro,
/// formControlMacro, worksheetFunction, dynamicDispatch. <c>Procedure</c> is the resolved FQN
/// (<c>Module.Proc</c>) or null when <c>Resolved</c> is false; <c>Target</c> keeps the raw reference.
/// </summary>
public sealed record ExcelVbaEntryPoint(
    string? Procedure,
    string Kind,
    string? Sheet,
    string? ShapeName,
    IReadOnlyList<string>? FormulaCells,
    string? Target,
    bool Resolved,
    ExcelVbaSiteRef? Site);

/// <summary>A procedure with no path from any entry point. <c>Confidence</c> is high or medium (see design doc).</summary>
public sealed record ExcelVbaUnreachableProcedure(
    string Procedure,
    string Module,
    string ModuleKind,
    string? Scope,
    int LineCount,
    string Confidence);
