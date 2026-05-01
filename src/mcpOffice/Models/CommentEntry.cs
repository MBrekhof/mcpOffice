namespace McpOffice.Models;

public sealed record CommentEntry(int Id, string Author, DateTime Date, string Text, string AnchorText);
