using System.Drawing;
using DevExpress.Office.Utils;
using DevExpress.XtraRichEdit.API.Native;
using Markdig;
using Markdig.Extensions.Tables;
using MdTable = Markdig.Extensions.Tables.Table;
using MdTableRow = Markdig.Extensions.Tables.TableRow;
using MdTableCell = Markdig.Extensions.Tables.TableCell;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace McpOffice.Services.Word;

internal static class MarkdownToDocxConverter
{
    private static readonly MarkdownPipeline Pipeline =
        new MarkdownPipelineBuilder().UsePipeTables().Build();

    public static void Apply(Document document, string markdown, string? baseDirectory)
    {
        var ast = Markdown.Parse(markdown, Pipeline);
        var ctx = new ConversionContext(document, baseDirectory);
        foreach (var block in ast)
            WriteBlock(ctx, block);
    }

    private sealed record ConversionContext(Document Document, string? BaseDirectory);

    private static void WriteBlock(ConversionContext ctx, Block block)
    {
        switch (block)
        {
            case HeadingBlock h:
                WriteHeading(ctx, h);
                break;
            case ParagraphBlock p:
                WriteParagraph(ctx, p);
                break;
            case ListBlock list:
                WriteList(ctx, list, level: 0);
                break;
            case QuoteBlock q:
                WriteQuote(ctx, q);
                break;
            case ThematicBreakBlock:
                WriteHorizontalRule(ctx);
                break;
            case FencedCodeBlock fenced:
                WriteCodeBlock(ctx, fenced.Lines.ToString());
                break;
            case CodeBlock code:
                WriteCodeBlock(ctx, code.Lines.ToString());
                break;
            case MdTable mdTable:
                WriteTable(ctx, mdTable);
                break;
            // Other block kinds added in subsequent tasks.
            default:
                // Unknown blocks silently skipped; Serilog warning attached in Task 21.
                break;
        }
    }

    private static void WriteHeading(ConversionContext ctx, HeadingBlock block)
    {
        var styleName = $"Heading {Math.Clamp(block.Level, 1, 6)}";
        EnsureParagraphStyle(ctx.Document, styleName);
        var para = AppendNewParagraph(ctx);
        para.Style = ctx.Document.ParagraphStyles[styleName];
        if (block.Inline is null) return;
        foreach (var inline in block.Inline)
            WriteInline(ctx, para, inline);
    }

    private static void EnsureParagraphStyle(Document doc, string styleName)
    {
        if (doc.ParagraphStyles[styleName] is not null) return;
        var s = doc.ParagraphStyles.CreateNew();
        s.Name = styleName;
        doc.ParagraphStyles.Add(s);
    }

    private static void WriteParagraph(ConversionContext ctx, ParagraphBlock block)
    {
        var para = AppendNewParagraph(ctx);
        if (block.Inline is null) return;
        foreach (var inline in block.Inline)
            WriteInline(ctx, para, inline);
    }

    private static void WriteQuote(ConversionContext ctx, QuoteBlock block)
    {
        foreach (var child in block)
        {
            if (child is ParagraphBlock p)
            {
                var para = AppendNewParagraph(ctx);
                // 0.25" expressed in DevExpress document units (1/300th of an inch).
                para.LeftIndent = Units.InchesToDocumentsF(0.25f);
                if (p.Inline is null) continue;
                foreach (var inline in p.Inline)
                    WriteInline(ctx, para, inline);
            }
        }
    }

    private static Paragraph AppendNewParagraph(ConversionContext ctx)
    {
        var doc = ctx.Document;
        // DevExpress Document has no InsertParagraph(DocumentPosition); follow the existing
        // project pattern (WordDocumentService.InsertParagraph) of inserting "\n".
        doc.InsertText(doc.Range.End, "\n");
        return doc.Paragraphs[doc.Paragraphs.Count - 1];
    }

    private static void WriteList(ConversionContext ctx, ListBlock list, int level)
    {
        var doc = ctx.Document;

        // Create the abstract numbering list from the appropriate template.
        var template = list.IsOrdered
            ? doc.AbstractNumberingLists.NumberedListTemplate
            : doc.AbstractNumberingLists.BulletedListTemplate;
        var abstractList = template.CreateNew();
        doc.AbstractNumberingLists.Add(abstractList);

        // Create the concrete numbering list that references the abstract one.
        var numberingList = doc.NumberingLists.Add(abstractList.Index);
        var listIndex = numberingList.Index;

        foreach (var item in list.OfType<ListItemBlock>())
        {
            foreach (var sub in item)
            {
                switch (sub)
                {
                    case ParagraphBlock p:
                    {
                        var para = AppendNewParagraph(ctx);
                        para.ListIndex = listIndex;
                        para.ListLevel = level;
                        if (p.Inline is null) break;
                        foreach (var inline in p.Inline)
                            WriteInline(ctx, para, inline);
                        break;
                    }
                    case ListBlock nested:
                        WriteList(ctx, nested, level + 1);
                        break;
                }
            }
        }
    }

    private static readonly Color CodeBackground = Color.FromArgb(0xF2, 0xF2, 0xF2);

    private static void WriteCodeBlock(ConversionContext ctx, string text)
    {
        var doc = ctx.Document;
        var lines = text.Replace("\r\n", "\n").Split('\n');
        foreach (var line in lines)
        {
            var para = AppendNewParagraph(ctx);
            para.LeftIndent = Units.InchesToDocumentsF(0.1f);
            if (line.Length == 0) continue;

            var insertedRange = doc.InsertText(para.Range.End, line);
            var props = doc.BeginUpdateCharacters(insertedRange);
            try
            {
                props.FontName = "Consolas";
                props.FontSize = 9f;
                props.BackColor = CodeBackground;
            }
            finally { doc.EndUpdateCharacters(props); }
        }
    }

    private static void WriteHorizontalRule(ConversionContext ctx)
    {
        var para = AppendNewParagraph(ctx);
        var props = ctx.Document.BeginUpdateParagraphs(para.Range);
        try
        {
            props.Borders.BottomBorder.LineStyle = BorderLineStyle.Single;
            props.Borders.BottomBorder.LineWidth = 0.5f;
        }
        finally { ctx.Document.EndUpdateParagraphs(props); }
    }

    private static readonly Color HeaderBackground = Color.FromArgb(0xF2, 0xF2, 0xF2);

    private static void WriteTable(ConversionContext ctx, MdTable table)
    {
        var doc = ctx.Document;
        var rows = table.OfType<MdTableRow>().ToList();
        if (rows.Count == 0) return;

        var colCount = rows.Max(r => r.Count);
        if (colCount == 0) return;

        var dxTable = doc.Tables.Create(doc.Range.End, rows.Count, colCount);

        for (int r = 0; r < rows.Count; r++)
        {
            var mdRow = rows[r];
            for (int c = 0; c < mdRow.Count; c++)
            {
                var mdCell = (MdTableCell)mdRow[c];
                var dxCell = dxTable.Rows[r].Cells[c];

                // Collect the plain text from all inlines in the cell.
                // Using InsertText at ContentRange.Start matches WordDocumentService.InsertTable convention.
                var cellText = CollectCellText(mdCell);
                DocumentRange? insertedRange = null;
                if (cellText.Length > 0)
                    insertedRange = doc.InsertText(dxCell.ContentRange.Start, cellText);

                if (mdRow.IsHeader)
                {
                    dxCell.BackgroundColor = HeaderBackground;
                    // Bold the inserted text range (not ContentRange, which may include
                    // the trailing paragraph mark with undefined Bold).
                    if (insertedRange is not null)
                    {
                        var props = doc.BeginUpdateCharacters(insertedRange);
                        try { props.Bold = true; }
                        finally { doc.EndUpdateCharacters(props); }
                    }
                }

                // Apply GFM column alignment (`:---` left, `:---:` center, `---:` right).
                if (table.ColumnDefinitions is { } cols && c < cols.Count && cols[c].Alignment is { } align)
                {
                    var pProps = doc.BeginUpdateParagraphs(dxCell.ContentRange);
                    try
                    {
                        pProps.Alignment = align switch
                        {
                            TableColumnAlign.Left   => ParagraphAlignment.Left,
                            TableColumnAlign.Center => ParagraphAlignment.Center,
                            TableColumnAlign.Right  => ParagraphAlignment.Right,
                            _                       => ParagraphAlignment.Left,
                        };
                    }
                    finally { doc.EndUpdateParagraphs(pProps); }
                }
            }
        }
    }

    /// <summary>
    /// Extracts the plain-text content of a Markdig table cell.
    /// Concatenates all literal inlines; other inline types will be handled
    /// in Phase C (emphasis/links) by replacing this helper with per-inline WriteInline calls.
    /// </summary>
    private static string CollectCellText(MdTableCell cell)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var child in cell)
        {
            if (child is ParagraphBlock p && p.Inline is not null)
            {
                foreach (var inline in p.Inline)
                {
                    if (inline is LiteralInline lit)
                        sb.Append(lit.Content.ToString());
                    else if (inline is LineBreakInline)
                        sb.Append(' ');
                }
            }
        }
        return sb.ToString();
    }

    private static void WriteInline(ConversionContext ctx, Paragraph para, Inline inline)
    {
        switch (inline)
        {
            case LiteralInline lit:
                ctx.Document.InsertText(para.Range.End, lit.Content.ToString());
                break;
            // Bold/italic/code/links etc. added in subsequent tasks.
        }
    }
}
