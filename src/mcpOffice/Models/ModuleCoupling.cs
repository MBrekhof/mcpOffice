namespace McpOffice.Models;

public sealed record ModuleCoupling(
    string Module,
    int Ca,
    int Ce,
    double Instability,
    int InternalEdges);
