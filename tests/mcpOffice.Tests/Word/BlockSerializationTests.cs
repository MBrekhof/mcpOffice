using McpOffice.Models;
using System.Text.Json;

namespace McpOffice.Tests.Word;

/// <summary>
/// Guards the JSON shape of the polymorphic <see cref="Block"/> hierarchy. Every other
/// Word test asserts against the objects, which is exactly why nobody noticed that
/// `word_read_structured` sent `{}` for every block.
/// </summary>
public class BlockSerializationTests
{
    [Fact]
    public void HeadingBlock_serializes_its_content_through_the_base_type()
    {
        var json = SerializeAsBlocks(new HeadingBlock(2, "Datawarehouse"));

        Assert.Contains("\"type\":\"heading\"", json);
        Assert.Contains("\"level\":2", json);
        Assert.Contains("\"text\":\"Datawarehouse\"", json);
    }

    [Fact]
    public void ParagraphBlock_serializes_its_runs_through_the_base_type()
    {
        var block = new ParagraphBlock([new Run("Bold bit", true, false, null)]);

        var json = SerializeAsBlocks(block);

        Assert.Contains("\"type\":\"paragraph\"", json);
        Assert.Contains("\"text\":\"Bold bit\"", json);
        Assert.Contains("\"bold\":true", json);
    }

    [Fact]
    public void StructuredDocument_blocks_are_never_empty_objects()
    {
        var document = new StructuredDocument(
            [new HeadingBlock(1, "Titel"), new ParagraphBlock([new Run("Tekst", false, false, null)])],
            [],
            [],
            new DocumentMetadata(null, null, null, null, null, null, null, 1, 1, 2));

        var json = JsonSerializer.Serialize(document, WireOptions);

        Assert.DoesNotContain("{}", json);
    }

    /// <summary>
    /// Serializes through <c>IReadOnlyList&lt;Block&gt;</c> — the declared type the service
    /// returns, and the one that dropped the properties — with the camelCase web defaults
    /// the MCP SDK uses, so these assertions describe the actual wire shape.
    /// </summary>
    private static string SerializeAsBlocks(Block block) =>
        JsonSerializer.Serialize<IReadOnlyList<Block>>([block], WireOptions);

    private static readonly JsonSerializerOptions WireOptions = new(JsonSerializerDefaults.Web);
}
