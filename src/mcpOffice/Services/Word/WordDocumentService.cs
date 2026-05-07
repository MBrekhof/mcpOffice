using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using MarkdownToDocxGenerator;
using McpOffice.Models;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using RichEditFormat = DevExpress.XtraRichEdit.DocumentFormat;

namespace McpOffice.Services.Word;

public sealed class WordDocumentService : IWordDocumentService
{
    public IReadOnlyList<OutlineNode> GetOutline(string path)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = new RichEditDocumentServer();
            server.LoadDocument(path, RichEditFormat.OpenXml);

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
            server.SaveDocument(path, RichEditFormat.OpenXml);
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    public string InsertParagraph(string path, int atIndex, string text, string? style)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var document = server.Document;
            var paragraphCount = document.Paragraphs.Count;

            if (atIndex < 0 || atIndex > paragraphCount)
            {
                throw ToolError.IndexOutOfRange(atIndex, paragraphCount);
            }

            var insertPos = atIndex == paragraphCount
                ? document.Range.End
                : document.Paragraphs[atIndex].Range.Start;

            var insertedRange = document.InsertText(insertPos, text + "\n");

            if (!string.IsNullOrEmpty(style))
            {
                EnsureParagraphStyle(document, style);
                var paragraph = document.Paragraphs.Get(insertedRange).First();
                paragraph.Style = document.ParagraphStyles[style];
            }

            server.SaveDocument(path, RichEditFormat.OpenXml);
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    private static readonly Regex MailMergeTokenPattern = new(@"\{\{(\w+)\}\}", RegexOptions.Compiled);

    public string MailMerge(string templatePath, string outputPath, string dataJson)
    {
        PathGuard.RequireExists(templatePath);
        PathGuard.RequireWritable(outputPath, overwrite: false);

        Dictionary<string, JsonElement> data;
        try
        {
            data = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(dataJson)
                   ?? new Dictionary<string, JsonElement>();
        }
        catch (JsonException ex)
        {
            throw ToolError.ParseError("dataJson", ex.Message);
        }

        try
        {
            using var server = LoadOpenXml(templatePath);
            var document = server.Document;
            var fullText = document.GetText(document.Range);

            var tokens = MailMergeTokenPattern.Matches(fullText)
                .Select(m => m.Groups[1].Value)
                .Distinct()
                .ToList();

            var missing = tokens.Where(t => !data.ContainsKey(t)).ToList();
            if (missing.Count > 0)
            {
                throw ToolError.MergeFieldMissing(missing);
            }

            foreach (var token in tokens)
            {
                var find = "{{" + token + "}}";
                var replacement = data[token].ValueKind == JsonValueKind.String
                    ? data[token].GetString() ?? string.Empty
                    : data[token].ToString();
                document.ReplaceAll(find, replacement, SearchOptions.None);
            }

            server.SaveDocument(outputPath, RichEditFormat.OpenXml);
            return outputPath;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    private static RichEditDocumentServer LoadInput(string inputPath)
    {
        var ext = Path.GetExtension(inputPath);
        if (ext.Equals(".md", StringComparison.OrdinalIgnoreCase) ||
            ext.Equals(".markdown", StringComparison.OrdinalIgnoreCase))
        {
            var server = new RichEditDocumentServer();
            var md = File.ReadAllText(inputPath, Encoding.UTF8);
            MarkdownToDocxConverter.Apply(server.Document, md, Path.GetDirectoryName(inputPath));
            return server;
        }
        return LoadOpenXml(inputPath);
    }

    public string Convert(string inputPath, string outputPath, string? format)
    {
        PathGuard.RequireExists(inputPath);
        PathGuard.RequireWritable(outputPath, overwrite: false);

        var outputFormat = ResolveOutputFormat(format, outputPath);

        try
        {
            using var server = LoadInput(inputPath);

            switch (outputFormat)
            {
                case WordOutputFormat.Pdf:
                    server.ExportToPdf(outputPath);
                    break;
                case WordOutputFormat.Html:
                    server.SaveDocument(outputPath, RichEditFormat.Html);
                    break;
                case WordOutputFormat.Rtf:
                    server.SaveDocument(outputPath, RichEditFormat.Rtf);
                    break;
                case WordOutputFormat.Text:
                    server.SaveDocument(outputPath, RichEditFormat.PlainText);
                    break;
                case WordOutputFormat.Markdown:
                    File.WriteAllText(outputPath, RenderMarkdown(server), Encoding.UTF8);
                    break;
                case WordOutputFormat.OpenXml:
                    server.SaveDocument(outputPath, RichEditFormat.OpenXml);
                    break;
                default:
                    throw ToolError.UnsupportedFormat(format ?? Path.GetExtension(outputPath));
            }

            return outputPath;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    public string SetMetadata(string path, IReadOnlyDictionary<string, string> properties)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var docProps = server.Document.DocumentProperties;

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "author":
                        docProps.Author = value;
                        break;
                    case "title":
                        docProps.Title = value;
                        break;
                    case "subject":
                        docProps.Subject = value;
                        break;
                    case "keywords":
                        docProps.Keywords = value;
                        break;
                    default:
                        throw ToolError.UnsupportedFormat(key);
                }
            }

            server.SaveDocument(path, RichEditFormat.OpenXml);
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    public string InsertTable(string path, int atIndex, IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var document = server.Document;
            var paragraphCount = document.Paragraphs.Count;

            if (atIndex < 0 || atIndex > paragraphCount)
            {
                throw ToolError.IndexOutOfRange(atIndex, paragraphCount);
            }

            var insertPos = atIndex == paragraphCount
                ? document.Range.End
                : document.Paragraphs[atIndex].Range.Start;

            var totalRows = 1 + rows.Count;
            var totalCols = headers.Count;
            if (totalCols == 0)
            {
                throw ToolError.ParseError(path, "headers must contain at least one column");
            }

            var table = document.Tables.Create(insertPos, totalRows, totalCols);

            for (var c = 0; c < headers.Count; c++)
            {
                document.InsertText(table.Rows[0].Cells[c].ContentRange.Start, headers[c]);
            }

            for (var r = 0; r < rows.Count; r++)
            {
                var rowCells = rows[r];
                for (var c = 0; c < rowCells.Count && c < totalCols; c++)
                {
                    document.InsertText(table.Rows[r + 1].Cells[c].ContentRange.Start, rowCells[c]);
                }
            }

            server.SaveDocument(path, RichEditFormat.OpenXml);
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    public ReplaceResult FindReplace(string path, string find, string replace, bool useRegex, bool matchCase)
    {
        PathGuard.RequireExists(path);

        try
        {
            using var server = LoadOpenXml(path);
            var document = server.Document;

            int count;
            if (useRegex)
            {
                var regexOptions = matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                count = document.ReplaceAll(new Regex(find, regexOptions), replace);
            }
            else
            {
                var options = matchCase
                    ? SearchOptions.CaseSensitive
                    : SearchOptions.None;
                count = document.ReplaceAll(find, replace, options);
            }

            server.SaveDocument(path, RichEditFormat.OpenXml);
            return new ReplaceResult(count);
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
            var baseDir = Path.GetDirectoryName(path);
            MarkdownToDocxConverter.Apply(server.Document, markdown ?? string.Empty, baseDir);
            server.SaveDocument(path, RichEditFormat.OpenXml);
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
            var baseDir = Path.GetDirectoryName(path);
            MarkdownToDocxConverter.Apply(server.Document, markdown ?? string.Empty, baseDir);
            server.SaveDocument(path, RichEditFormat.OpenXml);
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
            return RenderMarkdown(server);
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.ParseError(path, ex.Message);
        }
    }

    private enum WordOutputFormat
    {
        Pdf,
        Html,
        Rtf,
        Text,
        Markdown,
        OpenXml
    }

    private static WordOutputFormat ResolveOutputFormat(string? format, string outputPath)
    {
        var value = string.IsNullOrWhiteSpace(format)
            ? Path.GetExtension(outputPath).TrimStart('.')
            : format.Trim().TrimStart('.');

        return value.ToLowerInvariant() switch
        {
            "pdf" => WordOutputFormat.Pdf,
            "html" or "htm" => WordOutputFormat.Html,
            "rtf" => WordOutputFormat.Rtf,
            "txt" or "text" => WordOutputFormat.Text,
            "md" or "markdown" => WordOutputFormat.Markdown,
            "docx" => WordOutputFormat.OpenXml,
            _ => throw ToolError.UnsupportedFormat(value)
        };
    }

    private static string RenderMarkdown(RichEditDocumentServer server)
    {
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

    private static RichEditDocumentServer LoadOpenXml(string path)
    {
        var server = new RichEditDocumentServer();
        server.LoadDocument(path, RichEditFormat.OpenXml);
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

    private static RichEditDocumentServer CreateDocumentFromMarkdown(string? markdown)
    {
        var generator = new MdReportGenenerator(
            NullLogger<MdReportGenenerator>.Instance,
            new MdToOxmlEngine(NullLogger<MdToOxmlEngine>.Instance));

        using var stream = generator.TransformWithStream([markdown ?? string.Empty]);
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        var server = new RichEditDocumentServer();
        server.LoadDocument(stream, RichEditFormat.OpenXml);
        NormalizeMarkdownGeneratedDocument(server.Document, markdown);
        return server;
    }

    private sealed record MarkdownHeading(int Level, string Text);

    private static void NormalizeMarkdownGeneratedDocument(Document document, string? markdown)
    {
        ApplyMarkdownHeadingStyles(document, ExtractMarkdownHeadings(markdown));
        ApplyMarkdownItalicStyles(document, ExtractMarkdownItalicSpans(markdown));

        foreach (var paragraph in document.Paragraphs)
        {
            var styleName = paragraph.Style?.Name;
            var match = styleName is null
                ? System.Text.RegularExpressions.Match.Empty
                : Regex.Match(styleName, @"^Titre(?<level>[1-6])$", RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                continue;
            }

            var headingStyle = $"Heading {match.Groups["level"].Value}";
            EnsureParagraphStyle(document, headingStyle);
            paragraph.Style = document.ParagraphStyles[headingStyle];
        }
    }

    private static IReadOnlyList<MarkdownHeading> ExtractMarkdownHeadings(string? markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
        {
            return [];
        }

        var headings = new List<MarkdownHeading>();
        var normalized = markdown.Replace("\r\n", "\n").Replace("\r", "\n");
        foreach (var line in normalized.Split('\n'))
        {
            var match = Regex.Match(line, @"^\s*(#{1,6})\s+(.+?)\s*#*\s*$");
            if (!match.Success)
            {
                continue;
            }

            headings.Add(new MarkdownHeading(
                match.Groups[1].Value.Length,
                StripInlineMarkdown(match.Groups[2].Value.Trim())));
        }

        return headings;
    }

    private static void ApplyMarkdownHeadingStyles(Document document, IReadOnlyList<MarkdownHeading> headings)
    {
        if (headings.Count == 0)
        {
            return;
        }

        var headingIndex = 0;
        foreach (var paragraph in document.Paragraphs)
        {
            if (headingIndex >= headings.Count)
            {
                return;
            }

            var text = document.GetText(paragraph.Range).Trim();
            var heading = headings[headingIndex];
            if (!string.Equals(text, heading.Text, StringComparison.Ordinal))
            {
                continue;
            }

            var headingStyle = $"Heading {heading.Level}";
            EnsureParagraphStyle(document, headingStyle);
            paragraph.Style = document.ParagraphStyles[headingStyle];
            headingIndex++;
        }
    }

    private static string StripInlineMarkdown(string text)
    {
        var result = Regex.Replace(text, @"\*\*(.+?)\*\*", "$1");
        result = Regex.Replace(result, @"\*(.+?)\*", "$1");
        result = Regex.Replace(result, @"\[(.+?)\]\(.+?\)", "$1");
        return result;
    }

    private static IReadOnlyList<string> ExtractMarkdownItalicSpans(string? markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
        {
            return [];
        }

        var spans = new List<string>();
        var withoutCodeFences = Regex.Replace(markdown, @"```.*?```", string.Empty, RegexOptions.Singleline);
        foreach (System.Text.RegularExpressions.Match match in Regex.Matches(
                     withoutCodeFences,
                     @"(?<!\*)\*(?!\*)(?<text>.+?)(?<!\*)\*(?!\*)"))
        {
            var text = match.Groups["text"].Value.Trim();
            if (text.Length > 0)
            {
                spans.Add(StripInlineMarkdown(text));
            }
        }

        return spans.Distinct(StringComparer.Ordinal).ToList();
    }

    private static void ApplyMarkdownItalicStyles(Document document, IReadOnlyList<string> spans)
    {
        foreach (var span in spans)
        {
            foreach (var range in document.FindAll(span, SearchOptions.None))
            {
                var properties = document.BeginUpdateCharacters(range);
                properties.Italic = true;
                document.EndUpdateCharacters(properties);
            }
        }
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
