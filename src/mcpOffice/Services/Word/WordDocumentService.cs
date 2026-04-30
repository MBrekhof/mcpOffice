using DevExpress.XtraRichEdit;
using McpOffice.Models;
using ModelContextProtocol;
using System.Text;
using System.Text.RegularExpressions;

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

    public DocumentMetadata GetMetadata(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var document = server.Document;
            var properties = document.DocumentProperties;
            var text = document.GetText(document.Range);

            return new DocumentMetadata(
                EmptyToNull(properties.Author),
                EmptyToNull(properties.Title),
                EmptyToNull(properties.Subject),
                EmptyToNull(properties.Keywords),
                NullIfDefault(properties.Created),
                NullIfDefault(properties.Modified),
                NullIfDefault(properties.LastPrinted),
                properties.Revision,
                server.DocumentLayout.GetPageCount(),
                CountWords(text));
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public string ReadAsMarkdown(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var document = server.Document;
            var markdown = new StringBuilder();

            foreach (var paragraph in document.Paragraphs)
            {
                var text = document.GetText(paragraph.Range).Trim();
                if (text.Length == 0)
                {
                    continue;
                }

                var headingLevel = TryGetHeadingLevel(paragraph.Style?.Name);
                if (headingLevel is not null)
                {
                    markdown.Append('#', headingLevel.Value);
                    markdown.Append(' ');
                    markdown.AppendLine(text);
                    markdown.AppendLine();
                    continue;
                }

                markdown.AppendLine(text);
                markdown.AppendLine();
            }

            return markdown.ToString().TrimEnd();
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    private static RichEditDocumentServer LoadOpenXml(string path)
    {
        var server = new RichEditDocumentServer();
        server.LoadDocument(path, DocumentFormat.OpenXml);
        return server;
    }

    private static int? TryGetHeadingLevel(string? styleName)
    {
        const string prefix = "Heading ";

        if (styleName is null || !styleName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return int.TryParse(styleName[prefix.Length..], out var level) ? level : null;
    }

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private static DateTime? NullIfDefault(DateTime value) =>
        value == default ? null : value;

    private static int CountWords(string text) =>
        Regex.Matches(text, @"\b[\p{L}\p{N}]+(?:['-][\p{L}\p{N}]+)?\b").Count;
}
