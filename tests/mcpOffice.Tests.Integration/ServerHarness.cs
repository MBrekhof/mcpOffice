using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Client;

namespace McpOffice.Tests.Integration;

public sealed class ServerHarness : IAsyncDisposable
{
    public McpClient Client { get; private init; } = null!;

    public static async Task<ServerHarness> StartAsync()
    {
        var repoRoot = FindRepoRoot();
        var serverDll = Path.Combine(
            repoRoot,
            "src",
            "mcpOffice",
            "bin",
            "Debug",
            "net9.0",
            "mcpOffice.dll");

        if (!File.Exists(serverDll))
        {
            throw new FileNotFoundException(
                $"Server build output missing: {serverDll}. Run 'dotnet build' first.",
                serverDll);
        }

        var transport = new StdioClientTransport(
            new StdioClientTransportOptions
            {
                Name = "mcpOffice",
                Command = "dotnet",
                Arguments = [serverDll],
                WorkingDirectory = repoRoot
            },
            NullLoggerFactory.Instance);

        return new ServerHarness
        {
            Client = await McpClient.CreateAsync(transport)
        };
    }

    private static string FindRepoRoot()
    {
        var asmDir = Path.GetDirectoryName(typeof(ServerHarness).Assembly.Location)!;
        var dir = new DirectoryInfo(asmDir);

        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "mcpOffice.sln")))
        {
            dir = dir.Parent;
        }

        return dir?.FullName ?? throw new InvalidOperationException("Could not locate repo root.");
    }

    public async ValueTask DisposeAsync() => await Client.DisposeAsync();
}
