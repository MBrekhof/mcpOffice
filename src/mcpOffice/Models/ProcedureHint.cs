namespace McpOffice.Models;

public sealed record ProcedureHint(
    string Module,
    string ProcedureName,
    string Kind,
    bool IsEventHandler,
    int ParamCount,
    int CallerCount,
    int CalleeCount,
    ProcedureAxes Axes,
    string Rationale,
    CSharpSuggestion? CsharpSuggestion);
