namespace McpOffice.Models;

/// <summary>Result of <c>excel_compare_vba_corpus</c>: VBA procedures shared across several workbooks.</summary>
public sealed record ExcelVbaCorpusResult(
    IReadOnlyList<ExcelVbaCorpusWorkbook> Workbooks,
    ExcelVbaCorpusSummary Summary,
    IReadOnlyList<ExcelVbaSharedProcedure> SharedProcedures,
    IReadOnlyList<ExcelVbaSharedModule> SharedModules,
    bool Truncated);

public sealed record ExcelVbaCorpusWorkbook(
    string Path,
    bool HasVbaProject,
    int ModuleCount,
    int ProcedureCount,
    string? Error);

public sealed record ExcelVbaCorpusSummary(
    int WorkbookCount,
    int ProcedureCount,
    int SharedProcedureCount,
    int IdenticalGroups,
    int NearDuplicateGroups,
    int SharedModuleCount);

/// <summary><c>Tier</c> is identical or nearDuplicate; <c>Name</c> is the most common procedure name in the group.</summary>
public sealed record ExcelVbaSharedProcedure(
    string Tier,
    string Name,
    int LineCount,
    IReadOnlyList<ExcelVbaProcedureOccurrence> Occurrences);

public sealed record ExcelVbaProcedureOccurrence(
    string Workbook,
    string Module,
    string Procedure,
    double Similarity);

/// <summary>A module name that appears in several workbooks with mostly shared procedures.</summary>
public sealed record ExcelVbaSharedModule(
    string Module,
    IReadOnlyList<string> Workbooks,
    double SharedProcedureRatio);
