using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba.Rendering;

public sealed record CallgraphRenderOptions(
    string Layout = "clustered");

public interface ICallgraphRenderer
{
    string Render(FilteredCallgraph graph, CallgraphRenderOptions options);
}
