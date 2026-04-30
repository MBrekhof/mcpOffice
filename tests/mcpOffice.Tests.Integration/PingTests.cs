using ModelContextProtocol.Protocol;

namespace McpOffice.Tests.Integration;

public class PingTests
{
    [Fact]
    public async Task Lists_ping_tool()
    {
        await using var harness = await ServerHarness.StartAsync();

        var tools = await harness.Client.ListToolsAsync();

        Assert.Contains(tools, t => t.Name == "Ping");
    }

    [Fact]
    public async Task Ping_returns_pong()
    {
        await using var harness = await ServerHarness.StartAsync();

        var result = await harness.Client.CallToolAsync("Ping", new Dictionary<string, object?>());
        var text = result.Content.OfType<TextContentBlock>().Single().Text;

        Assert.Equal("pong", text);
    }
}
