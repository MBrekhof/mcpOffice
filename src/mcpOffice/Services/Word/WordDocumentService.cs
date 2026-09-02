using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using McpOffice.Models;
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
            using var server = Load(path);

            var document = server.Document;
            var outline = new List<OutlineNode>();

            foreach (var paragraph in document.Paragraphs)
            {
                var level = TryGetHeadingLevel(paragraph);
                if (level is null)
                {
                    continue;
                }

                var text = StripHeadingNumberLabel(document, paragraph, document.GetText(paragraph.Range).Trim());
                if (text.Length > 0)
                {
                    outline.Add(new OutlineNode(level.Value, text));
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
            using var server = Load(path);
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
            using var server = Load(path);
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

                var headingLevel = TryGetHeadingLevel(paragraph);
                if (headingLevel is not null)
                {
                    blocks.Add(new HeadingBlock(headingLevel.Value, StripHeadingNumberLabel(document, paragraph, text)));
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
            using var server = Load(path);
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
            using var server = Load(path);
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
            server.SaveDocument(path, WordFormats.ForPath(path));
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
            using var server = Load(path);
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

            server.SaveDocument(path, WordFormats.ForPath(path));
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    private static readonly Regex MailMergeTokenPattern = new(@"\{\{(\w+)\}\}", RegexOptions.Compiled);

    public string MailMerge(string templatePath, string outputPath, string dataJson, bool overwrite = false)
    {
        PathGuard.RequireExists(templatePath);
        PathGuard.RequireWritable(outputPath, overwrite);

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
            using var server = Load(templatePath);
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

            server.SaveDocument(outputPath, WordFormats.ForPath(outputPath));
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
        return Load(inputPath);
    }

    public string Convert(string inputPath, string outputPath, string? format, bool overwrite = false)
    {
        PathGuard.RequireExists(inputPath);
        PathGuard.RequireWritable(outputPath, overwrite);

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
                case WordOutputFormat.OpenDocument:
                    server.SaveDocument(outputPath, RichEditFormat.Odt);
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
            using var server = Load(path);
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

            server.SaveDocument(path, WordFormats.ForPath(path));
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
            using var server = Load(path);
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

            server.SaveDocument(path, WordFormats.ForPath(path));
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
            using var server = Load(path);
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

            server.SaveDocument(path, WordFormats.ForPath(path));
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
            using var server = Load(path);
            var baseDir = Path.GetDirectoryName(path);
            MarkdownToDocxConverter.Apply(server.Document, markdown ?? string.Empty, baseDir);
            server.SaveDocument(path, WordFormats.ForPath(path));
            return path;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            throw ToolError.IoError(ex.Message);
        }
    }

    public string CreateFromMarkdown(string path, string markdown, bool overwrite, string? templatePath = null)
    {
        PathGuard.RequireWritable(path, overwrite);
        if (templatePath is not null)
        {
            PathGuard.RequireExists(templatePath);
        }

        try
        {
            using var server = new RichEditDocumentServer();
            if (templatePath is not null)
            {
                server.LoadDocumentTemplate(templatePath);
            }
            var baseDir = Path.GetDirectoryName(path);
            MarkdownToDocxConverter.Apply(
                server.Document,
                markdown ?? string.Empty,
                baseDir,
                preserveExistingHeadingStyles: templatePath is not null);
            server.SaveDocument(path, WordFormats.ForPath(path));
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
            using var server = Load(path);
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
        OpenXml,
        OpenDocument
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
            "odt" => WordOutputFormat.OpenDocument,
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

            var headingLevel = TryGetHeadingLevel(paragraph);
            if (headingLevel is not null)
            {
                markdown.Append('#', headingLevel.Value);
                markdown.Append(' ');
                markdown.AppendLine(StripHeadingNumberLabel(document, paragraph, text));
                markdown.AppendLine();
                continue;
            }

            markdown.AppendLine(text);
            markdown.AppendLine();
        }

        return markdown.ToString().TrimEnd();
    }

    /// <summary>
    /// Loads a document in the format its extension implies (see <see cref="WordFormats"/>).
    /// Anything unrecognised is read as OpenXML.
    /// </summary>
    private static RichEditDocumentServer Load(string path)
    {
        var server = new RichEditDocumentServer();
        server.LoadDocument(path, WordFormats.ForPath(path));
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

    /// <summary>
    /// "Heading 1" (OpenXML) and "Heading1" (what the ODT importer produces) both count.
    /// </summary>
    private static readonly Regex HeadingStyleNamePattern =
        new(@"^heading\s*([1-9])$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static int? TryGetHeadingLevel(string? styleName)
    {
        if (styleName is null)
        {
            return null;
        }

        var match = HeadingStyleNamePattern.Match(styleName.Trim());
        return match.Success ? int.Parse(match.Groups[1].Value) : null;
    }

    /// <summary>
    /// A "{0}" placeholder inside a <see cref="Regex.Escape(string)"/>d format string.
    /// Escape turns "{" into "\{" but leaves "}" alone, hence the optional backslashes.
    /// </summary>
    private static readonly Regex ListLabelPlaceholderPattern = new(@"\\?\{\d+\\?\}", RegexOptions.Compiled);

    /// <summary>
    /// Drops the numbering label from a numbered heading's text.
    /// <para>
    /// <c>Document.GetText</c> renders the list label into the paragraph text, and for a
    /// document that came in through ODT the counters are not resolved — every level comes
    /// out as "1", so section 1.2 reads "1.1.Versiebeheer". A wrong section number is worse
    /// for an agent than no number: one citing "section 1.1" would be citing nothing.
    /// </para>
    /// <para>
    /// The label to remove is derived from the list level's own
    /// <c>DisplayFormatString</c> ("{0}.{1}." and the like) rather than guessed, so exactly
    /// the rendered label goes and a number the author typed into the heading survives.
    /// </para>
    /// </summary>
    private static string StripHeadingNumberLabel(Document document, Paragraph paragraph, string text)
    {
        var pattern = TryGetListLabelPattern(document, paragraph);
        if (pattern is null)
        {
            return text;
        }

        var match = pattern.Match(text);
        if (!match.Success || match.Length == 0)
        {
            return text;
        }

        var stripped = text[match.Length..];
        return stripped.Length > 0 ? stripped : text;
    }

    /// <summary>
    /// Turns the paragraph's list level format ("{0}.{1}.") into a regex anchored at the
    /// start of the text (<c>^\s*\d+\.\d+\.\s*</c>). Null when the paragraph is not in a
    /// numbering list, or the level has no numeric format to strip.
    /// </summary>
    private static Regex? TryGetListLabelPattern(Document document, Paragraph paragraph)
    {
        if (paragraph.ListIndex < 0 || paragraph.ListLevel is < 0 or > 8)
        {
            return null;
        }

        string? format;
        try
        {
            format = document.NumberingLists[paragraph.ListIndex]
                .Levels[paragraph.ListLevel]
                .DisplayFormatString;
        }
        catch (Exception ex) when (ex is not McpException)
        {
            return null;
        }

        if (string.IsNullOrEmpty(format) || !format.Contains('{'))
        {
            return null;
        }

        var body = ListLabelPlaceholderPattern.Replace(Regex.Escape(format), @"\d+");
        return new Regex($@"^\s*{body}\s*");
    }

    /// <summary>
    /// Heading level from the style name, falling back to the paragraph's outline level.
    /// The fallback is what makes a document with renamed or localised heading styles
    /// readable — an ODT round-trip through Word keeps the outline level even where the
    /// style name is document-specific (e.g. "Hoofdstkbijlagen"). 0 means body text.
    /// </summary>
    private static int? TryGetHeadingLevel(Paragraph paragraph)
    {
        var fromStyleName = TryGetHeadingLevel(paragraph.Style?.Name);
        if (fromStyleName is not null)
        {
            return fromStyleName;
        }

        var outlineLevel = paragraph.OutlineLevel;
        return outlineLevel is >= 1 and <= 9 ? outlineLevel : null;
    }

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private static DateTime? NullIfDefault(DateTime value) =>
        value == default ? null : value;

    private static int CountWords(string text) =>
        Regex.Matches(text, @"\b[\p{L}\p{N}]+(?:['-][\p{L}\p{N}]+)?\b").Count;
}
