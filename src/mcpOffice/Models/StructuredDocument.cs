namespace McpOffice.Models;

public sealed record StructuredDocument(
    IReadOnlyList<Block> Blocks,
    IReadOnlyList<TableBlock> Tables,
    IReadOnlyList<ImageRef> Images,
    DocumentMetadata Properties);
