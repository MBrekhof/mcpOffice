using DevExpress.XtraRichEdit;
using McpOffice.Models;
using ModelContextProtocol;

namespace McpOffice.Services.Word;

public sealed class WordDocumentService : IWordDocumentService
{
    public IReadOnlyList<OutlineNode> GetOutline(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = new RichEditDocumentServer();
            server.LoadDocument(path, DocumentFormat.OpenXml);

            var document = server.Document;
            var outline = new List<OutlineNode>();

            foreach (var paragraph in document.Paragraphs)
            {
                var styleName = paragraph.Style?.Name ?? string.Empty;
                if (!styleName.StartsWith("Heading ", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!int.TryParse(styleName["Heading ".Length..], out var level))
                {
                    continue;
                }

                var text = document.GetText(paragraph.Range).Trim();
                if (text.Length > 0)
                {
                    outline.Add(new OutlineNode(level, text));
                }
            }

            return outline;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }
}
