using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

/// <summary>
/// Gated benchmark against a real Word-exported .odt (1.7 MB, Dutch-language manual with
/// headings, numbered lists, tables and images). Skips when the file is absent, like
/// <c>Excel/Vba/AirSampleAnalysisTests</c>. The bar is meaning, not styling: an agent must
/// be able to answer questions about the document from what these tools return.
/// </summary>
public class OdtRealWorldTests
{
    private const string SamplePath =
        @"C:\Projects\WLNCentral\rewab\20221220 Handleiding Risicogestuurd monitoren.odt";

    private static bool SampleMissing => !File.Exists(SamplePath);

    [Fact]
    public void Real_word_exported_odt_yields_an_outline()
    {
        if (SampleMissing)
        {
            return;
        }

        var nodes = new WordDocumentService().GetOutline(SamplePath);

        Assert.NotEmpty(nodes);
        Assert.All(nodes, node => Assert.False(string.IsNullOrWhiteSpace(node.Text)));
        // No heading in this document is numbered by hand, so any leading "1." would be
        // the unresolved label the ODT import renders into the text.
        Assert.All(nodes, node => Assert.DoesNotMatch(@"^\d+(\.\d+)*\.", node.Text));
    }

    [Fact]
    public void Real_word_exported_odt_yields_readable_markdown()
    {
        if (SampleMissing)
        {
            return;
        }

        var markdown = new WordDocumentService().ReadAsMarkdown(SamplePath);

        Assert.True(markdown.Length > 1000, $"expected substantial text, got {markdown.Length} chars");
        Assert.Contains("# ", markdown);
    }

    [Fact]
    public void Real_word_exported_odt_yields_structured_blocks_and_tables()
    {
        if (SampleMissing)
        {
            return;
        }

        var structured = new WordDocumentService().ReadStructured(SamplePath);

        Assert.NotEmpty(structured.Blocks);
        Assert.All(
            structured.Tables,
            table => Assert.All(table.Rows, row => Assert.NotEmpty(row)));
    }
}
