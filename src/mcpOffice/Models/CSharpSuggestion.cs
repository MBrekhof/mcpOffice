namespace McpOffice.Models;

public sealed record CSharpSuggestion(
    string TargetType,
    string SuggestedClassName,
    string SuggestedMethodName,
    string? Lifetime,
    bool IsPublic,
    IReadOnlyList<string> Blockers);
