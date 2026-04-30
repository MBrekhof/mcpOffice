using System.ComponentModel;
using ModelContextProtocol.Server;

namespace McpOffice.Tools;

[McpServerToolType]
public static class PingTools
{
    [McpServerTool(Name = "Ping"), Description("Returns 'pong'. Use to verify the server is reachable.")]
    public static string Ping() => "pong";
}
