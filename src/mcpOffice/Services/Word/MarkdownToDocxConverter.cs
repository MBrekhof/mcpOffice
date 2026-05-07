using DevExpress.XtraRichEdit.API.Native;
using Markdig;
using Markdig.Syntax;

namespace McpOffice.Services.Word;

internal static class MarkdownToDocxConverter
{
    private static readonly MarkdownPipeline Pipeline =
        new MarkdownPipelineBuilder().UsePipeTables().Build();

    public static void Apply(Document document, string markdown, string? baseDirectory)
    {
        var ast = Markdown.Parse(markdown, Pipeline);
        // Block dispatch added in subsequent tasks.
        _ = ast;
        _ = document;
        _ = baseDirectory;
    }
}
