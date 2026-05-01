namespace McpOffice.Models;

public sealed record RevisionEntry(string Type, string Author, DateTime Date, string Text);
