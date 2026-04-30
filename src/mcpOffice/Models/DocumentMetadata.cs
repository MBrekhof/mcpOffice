namespace McpOffice.Models;

public sealed record DocumentMetadata(
    string? Author,
    string? Title,
    string? Subject,
    string? Keywords,
    DateTime? Created,
    DateTime? Modified,
    DateTime? LastPrinted,
    int RevisionCount,
    int PageCount,
    int WordCount);
