using McpOffice.Models;

namespace McpOffice.Services.Word;

public interface IWordDocumentService
{
    IReadOnlyList<OutlineNode> GetOutline(string path);
    DocumentMetadata GetMetadata(string path);
    string ReadAsMarkdown(string path);
}
