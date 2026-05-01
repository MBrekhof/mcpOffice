namespace McpOffice.Tests.Integration;

public class ToolSurfaceTests
{
    [Fact]
    public async Task Exposes_initial_tool_catalog()
    {
        string[] expected =
        [
            "Ping",
            "word_append_markdown",
            "word_create_blank",
            "word_create_from_markdown",
            "word_find_replace",
            "word_get_metadata",
            "word_get_outline",
            "word_insert_paragraph",
            "word_insert_table",
            "word_list_comments",
            "word_list_revisions",
            "word_read_markdown",
            "word_read_structured"
        ];

        await using var harness = await ServerHarness.StartAsync();
        var tools = await harness.Client.ListToolsAsync();
        var toolNames = tools.Select(t => t.Name).ToHashSet();

        Assert.Equal(expected.Order(), toolNames.Order());
    }
}
