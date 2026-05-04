using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba.Rendering;

public sealed record CallgraphRenderOptions(
    string Layout = "clustered");   // "clustered" | "flat"

public interface ICallgraphRenderer
{
    string Render(FilteredCallgraph graph, CallgraphRenderOptions options);
}
