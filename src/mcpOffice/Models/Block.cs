namespace McpOffice.Models;

public abstract record Block;

public sealed record HeadingBlock(int Level, string Text) : Block;

public sealed record ParagraphBlock(IReadOnlyList<Run> Runs) : Block;
