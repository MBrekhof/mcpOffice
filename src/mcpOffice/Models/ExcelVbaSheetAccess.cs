namespace McpOffice.Models;

/// <summary>Result of <c>excel_map_vba_sheet_access</c>: which sheet cells each procedure reads and writes.</summary>
public sealed record ExcelVbaSheetAccessResult(
    string Path,
    bool HasVbaProject,
    ExcelVbaSheetAccessSummary Summary,
    IReadOnlyList<ExcelVbaSheetAccess> SheetAccess,
    IReadOnlyList<ExcelVbaSheetUsage> Sheets,
    bool Truncated);

public sealed record ExcelVbaSheetAccessSummary(
    int SiteCount,
    int ResolvedCount,
    int UnresolvedCount,
    int SheetCount,
    int ProcedureCount);

public sealed record ExcelVbaSheetRef(string Name, string? CodeName);

/// <summary><c>Kind</c> is a closed set: range, definedName, column, row, wholeSheet, dynamicCells.</summary>
public sealed record ExcelVbaAccessTarget(string Kind, string? Address, string? DefinedName);

/// <summary>
/// One (procedure, sheet, target, mode) group. <c>Sheet</c> is null when the site could not be
/// attributed; <c>UnresolvedReason</c> then says why (activeSheet, aliasReassigned, unknownSheet,
/// unknownName). <c>Mode</c> is read, write or both.
/// </summary>
public sealed record ExcelVbaSheetAccess(
    string Procedure,
    ExcelVbaSheetRef? Sheet,
    ExcelVbaAccessTarget Target,
    string Mode,
    int SiteCount,
    string? UnresolvedReason);

/// <summary>Per-sheet rollup: who reads it, who writes it.</summary>
public sealed record ExcelVbaSheetUsage(
    string Name,
    string? CodeName,
    IReadOnlyList<string> Readers,
    IReadOnlyList<string> Writers,
    int ReadSites,
    int WriteSites);
