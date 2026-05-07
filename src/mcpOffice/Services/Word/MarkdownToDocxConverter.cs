using DevExpress.XtraRichEdit.API.Native;
using Markdig;
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

    private static Paragraph AppendNewParagraph(ConversionContext ctx)
    {
        var doc = ctx.Document;
        // DevExpress Document has no InsertParagraph(DocumentPosition); follow the existing
        // project pattern (WordDocumentService.InsertParagraph) of inserting "\n".
        doc.InsertText(doc.Range.End, "\n");
        return doc.Paragraphs[doc.Paragraphs.Count - 1];
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
