namespace McpOffice.Tests.Integration;

public class ToolSurfaceTests
{
    [Fact]
    public async Task Exposes_initial_tool_catalog()
    {
        string[] expected =
        [
            "Ping",
            "word_get_metadata",
            "word_get_outline",
            "word_read_markdown"
        ];

        await using var harness = await ServerHarness.StartAsync();
        var tools = await harness.Client.ListToolsAsync();
        var toolNames = tools.Select(t => t.Name).ToHashSet();

        Assert.Equal(expected.Order(), toolNames.Order());
    }
}
