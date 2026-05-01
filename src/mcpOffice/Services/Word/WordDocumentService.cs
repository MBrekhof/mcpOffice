using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
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
            return BuildMetadata(server);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public StructuredDocument ReadStructured(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var document = server.Document;

            var tableRanges = document.Tables
                .Select(t => (Start: t.Range.Start.ToInt(), End: t.Range.End.ToInt()))
                .ToList();

            var blocks = new List<Block>();
            foreach (var paragraph in document.Paragraphs)
            {
                var paraStart = paragraph.Range.Start.ToInt();
                if (tableRanges.Any(r => paraStart >= r.Start && paraStart < r.End))
                {
                    continue;
                }

                var rawText = document.GetText(paragraph.Range);
                var text = rawText.TrimEnd('\r', '\n', '\v', '\f');
                if (text.Length == 0)
                {
                    continue;
                }

                var headingLevel = TryGetHeadingLevel(paragraph.Style?.Name);
                if (headingLevel is not null)
                {
                    blocks.Add(new HeadingBlock(headingLevel.Value, text));
                    continue;
                }

                blocks.Add(new ParagraphBlock(BuildRuns(document, paragraph.Range.Start.ToInt(), text)));
            }

            var tables = new List<TableBlock>(document.Tables.Count);
            for (var i = 0; i < document.Tables.Count; i++)
            {
                var table = document.Tables[i];
                var rows = new List<IReadOnlyList<string>>(table.Rows.Count);
                for (var r = 0; r < table.Rows.Count; r++)
                {
                    var row = table.Rows[r];
                    var cells = new List<string>(row.Cells.Count);
                    for (var c = 0; c < row.Cells.Count; c++)
                    {
                        var cellText = document.GetText(row.Cells[c].ContentRange)
                            .TrimEnd('\r', '\n', '\v', '\f', '');
                        cells.Add(cellText);
                    }
                    rows.Add(cells);
                }
                tables.Add(new TableBlock(i, rows));
            }

            var properties = BuildMetadata(server);
            return new StructuredDocument(blocks, tables, Array.Empty<ImageRef>(), properties);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public IReadOnlyList<CommentEntry> ListComments(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var document = server.Document;
            var entries = new List<CommentEntry>(document.Comments.Count);

            for (var i = 0; i < document.Comments.Count; i++)
            {
                var comment = document.Comments[i];
                var anchorText = document.GetText(comment.Range);

                var body = comment.BeginUpdate();
                var commentText = body.GetText(body.Range).TrimEnd('\r', '\n', '\v', '\f');
                comment.EndUpdate(body);

                entries.Add(new CommentEntry(
                    i,
                    comment.Author ?? string.Empty,
                    comment.Date,
                    commentText,
                    anchorText));
            }

            return entries;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public IReadOnlyList<RevisionEntry> ListRevisions(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var document = server.Document;
            var entries = new List<RevisionEntry>();

            foreach (var revision in document.Revisions)
            {
                var text = document.GetText(revision.Range)
                    .TrimEnd('\r', '\n', '\v', '\f');
                entries.Add(new RevisionEntry(
                    MapRevisionType(revision.Type),
                    revision.Author ?? string.Empty,
                    revision.DateTime ?? default,
                    text));
            }

            return entries;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    public string CreateBlank(string path, bool overwrite)
    {
        PathGuard.RequireWritable(path, overwrite);

        try
        {
            using var server = new RichEditDocumentServer();
            server.SaveDocument(path, DocumentFormat.OpenXml);
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    public string AppendMarkdown(string path, string markdown)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            WriteMarkdownToDocument(server.Document, markdown);
            server.SaveDocument(path, DocumentFormat.OpenXml);
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    public string CreateFromMarkdown(string path, string markdown, bool overwrite)
    {
        PathGuard.RequireWritable(path, overwrite);

        try
        {
            using var server = new RichEditDocumentServer();
            WriteMarkdownToDocument(server.Document, markdown);
            server.SaveDocument(path, DocumentFormat.OpenXml);
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
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

    private static DocumentMetadata BuildMetadata(RichEditDocumentServer server)
    {
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

    private static IReadOnlyList<Run> BuildRuns(Document document, int paragraphStart, string text)
    {
        var runs = new List<Run>();
        if (text.Length == 0)
        {
            return runs;
        }

        var sb = new StringBuilder();
        bool? currentBold = null;
        bool? currentItalic = null;

        for (var i = 0; i < text.Length; i++)
        {
            var charRange = document.CreateRange(paragraphStart + i, 1);
            var props = document.BeginUpdateCharacters(charRange);
            var bold = props.Bold == true;
            var italic = props.Italic == true;
            document.EndUpdateCharacters(props);

            if (currentBold is null)
            {
                currentBold = bold;
                currentItalic = italic;
            }
            else if (currentBold != bold || currentItalic != italic)
            {
                runs.Add(new Run(sb.ToString(), currentBold ?? false, currentItalic ?? false, null));
                sb.Clear();
                currentBold = bold;
                currentItalic = italic;
            }

            sb.Append(text[i]);
        }

        if (sb.Length > 0)
        {
            runs.Add(new Run(sb.ToString(), currentBold ?? false, currentItalic ?? false, null));
        }

        return runs;
    }

    private static void WriteMarkdownToDocument(Document doc, string? markdown)
    {
        if (string.IsNullOrEmpty(markdown)) return;

        var normalized = markdown.Replace("\r\n", "\n").Replace("\r", "\n");
        var blocks = Regex.Split(normalized, @"\n\s*\n");

        foreach (var rawBlock in blocks)
        {
            var block = rawBlock.Trim();
            if (block.Length == 0) continue;

            var headingMatch = Regex.Match(block, @"^(#{1,6})\s+(.*)$", RegexOptions.Singleline);
            if (headingMatch.Success)
            {
                var level = headingMatch.Groups[1].Value.Length;
                var text = headingMatch.Groups[2].Value;
                var range = doc.AppendText(text + "\n");
                var paragraph = doc.Paragraphs.Get(range).First();
                EnsureParagraphStyle(doc, $"Heading {level}");
                paragraph.Style = doc.ParagraphStyles[$"Heading {level}"];
            }
            else
            {
                WriteInline(doc, block);
                doc.AppendText("\n");
            }
        }
    }

    private static void WriteInline(Document doc, string text)
    {
        var i = 0;
        while (i < text.Length)
        {
            if (i + 1 < text.Length && text[i] == '*' && text[i + 1] == '*')
            {
                var end = text.IndexOf("**", i + 2, StringComparison.Ordinal);
                if (end >= 0)
                {
                    AppendStyled(doc, text.Substring(i + 2, end - (i + 2)), bold: true, italic: false);
                    i = end + 2;
                    continue;
                }
            }
            if (text[i] == '*')
            {
                var end = text.IndexOf('*', i + 1);
                if (end >= 0)
                {
                    AppendStyled(doc, text.Substring(i + 1, end - (i + 1)), bold: false, italic: true);
                    i = end + 1;
                    continue;
                }
            }
            var nextSpecial = text.IndexOf('*', i);
            if (nextSpecial < 0) nextSpecial = text.Length;
            if (nextSpecial > i) doc.AppendText(text.Substring(i, nextSpecial - i));
            i = nextSpecial;
        }
    }

    private static void AppendStyled(Document doc, string text, bool bold, bool italic)
    {
        var range = doc.AppendText(text);
        var props = doc.BeginUpdateCharacters(range);
        if (bold) props.Bold = true;
        if (italic) props.Italic = true;
        doc.EndUpdateCharacters(props);
    }

    private static void EnsureParagraphStyle(Document doc, string styleName)
    {
        if (doc.ParagraphStyles[styleName] is not null) return;
        var style = doc.ParagraphStyles.CreateNew();
        style.Name = styleName;
        doc.ParagraphStyles.Add(style);
    }

    private static string MapRevisionType(RevisionType type) => type switch
    {
        RevisionType.Inserted => "insert",
        RevisionType.Deleted => "delete",
        RevisionType.CharacterPropertyChanged => "format",
        RevisionType.ParagraphPropertyChanged => "format",
        RevisionType.SectionPropertyChanged => "format",
        RevisionType.TablePropertyChanged => "format",
        RevisionType.TableRowPropertyChanged => "format",
        RevisionType.TableCellPropertyChanged => "format",
        RevisionType.CharacterStyleDefinitionChanged => "format",
        RevisionType.ParagraphStyleDefinitionChanged => "format",
        _ => type.ToString().ToLowerInvariant()
    };

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
