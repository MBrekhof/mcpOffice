using DevExpress.XtraRichEdit.API.Native;
using McpOffice.Services.Word;
using ModelContextProtocol;
using RichEditFormat = DevExpress.XtraRichEdit.DocumentFormat;

namespace McpOffice.Tests.Word;

/// <summary>
/// ODT is a read-first format here: the point is that an agent can get the *meaning*
/// of a Word-exported .odt out — headings, paragraph text, table cell text — not that
/// styling round-trips.
/// </summary>
public class OdtTests
{
    [Fact]
    public void Outline_returns_heading_tree_from_odt()
    {
        var path = CreateOdtDocument();
        var service = new WordDocumentService();

        var nodes = service.GetOutline(path);

        Assert.Collection(
            nodes,
            node =>
            {
                Assert.Equal(1, node.Level);
                Assert.Equal("Introduction", node.Text);
            },
            node =>
            {
                Assert.Equal(2, node.Level);
                Assert.Equal("Background", node.Text);
            });
    }

    [Fact]
    public void ReadAsMarkdown_returns_odt_content()
    {
        var path = CreateOdtDocument();
        var service = new WordDocumentService();

        var markdown = service.ReadAsMarkdown(path);

        Assert.Contains("# Introduction", markdown);
        Assert.Contains("## Background", markdown);
        Assert.Contains("Plain paragraph", markdown);
    }

    [Fact]
    public void ReadStructured_returns_odt_blocks_and_table_cells()
    {
        var path = TestWordDocuments.CreateOdt(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Title", "Heading 1");
            TestWordDocuments.AppendParagraph(document, "Body text");
            var table = document.Tables.Create(document.Range.End, 2, 2);
            document.InsertText(table.Rows[0].Cells[0].ContentRange.Start, "A1");
            document.InsertText(table.Rows[0].Cells[1].ContentRange.Start, "B1");
            document.InsertText(table.Rows[1].Cells[0].ContentRange.Start, "A2");
            document.InsertText(table.Rows[1].Cells[1].ContentRange.Start, "B2");
        });
        var service = new WordDocumentService();

        var structured = service.ReadStructured(path);

        Assert.Contains(structured.Blocks, b => b is McpOffice.Models.HeadingBlock { Text: "Title" });
        var table = Assert.Single(structured.Tables);
        Assert.Equal(new[] { "A1", "B1" }, table.Rows[0]);
        Assert.Equal(new[] { "A2", "B2" }, table.Rows[1]);
    }

    [Fact]
    public void GetMetadata_reads_odt()
    {
        var path = CreateOdtDocument();
        var service = new WordDocumentService();

        var metadata = service.GetMetadata(path);

        Assert.True(metadata.WordCount > 0);
    }

    [Fact]
    public void Convert_odt_to_markdown_keeps_content()
    {
        var path = CreateOdtDocument();
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.md");
        var service = new WordDocumentService();

        service.Convert(path, output, format: null);

        var markdown = File.ReadAllText(output);
        Assert.Contains("# Introduction", markdown);
        Assert.Contains("Plain paragraph", markdown);
    }

    [Fact]
    public void Convert_odt_to_pdf_produces_pdf()
    {
        var path = CreateOdtDocument();
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.pdf");
        var service = new WordDocumentService();

        service.Convert(path, output, format: null);

        var header = new byte[5];
        using (var stream = File.OpenRead(output))
        {
            stream.ReadExactly(header);
        }
        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(header));
    }

    [Fact]
    public void Convert_docx_to_odt_produces_readable_odt()
    {
        var path = TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "From docx", "Heading 1");
        });
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.odt");
        var service = new WordDocumentService();

        service.Convert(path, output, format: null);

        var nodes = service.GetOutline(output);
        Assert.Equal("From docx", Assert.Single(nodes).Text);
    }

    [Fact]
    public void FindReplace_on_odt_saves_back_as_odt()
    {
        var path = TestWordDocuments.CreateOdt(document =>
        {
            TestWordDocuments.AppendParagraph(document, "hello hello", "Heading 1");
        });
        var service = new WordDocumentService();

        var result = service.FindReplace(path, "hello", "hi", useRegex: false, matchCase: false);

        Assert.Equal(2, result.Replacements);
        // Re-reading proves the file is still a valid .odt, not OpenXML bytes under an .odt name.
        Assert.Contains("hi hi", service.ReadAsMarkdown(path));
    }

    [Fact]
    public void Outline_detects_space_less_heading_style_names()
    {
        // DevExpress's ODT import names heading styles "Heading1", not "Heading 1".
        var path = TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Kop een", "Heading1");
            TestWordDocuments.AppendParagraph(document, "Kop twee", "Heading2");
        });

        var nodes = new WordDocumentService().GetOutline(path);

        Assert.Collection(
            nodes,
            node => Assert.Equal((1, "Kop een"), (node.Level, node.Text)),
            node => Assert.Equal((2, "Kop twee"), (node.Level, node.Text)));
    }

    [Fact]
    public void Outline_falls_back_to_paragraph_outline_level()
    {
        // A heading whose style carries a document-specific name is still a heading if
        // Word gave it an outline level — that is what drives the navigation pane.
        var path = TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Bijlage A", "Hoofdstkbijlagen");
            document.Paragraphs[0].OutlineLevel = 1;
            TestWordDocuments.AppendParagraph(document, "Body text");
        });

        var nodes = new WordDocumentService().GetOutline(path);

        var node = Assert.Single(nodes);
        Assert.Equal(1, node.Level);
        Assert.Equal("Bijlage A", node.Text);
    }

    [Fact]
    public void ReadAsMarkdown_uses_outline_level_headings()
    {
        var path = TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Bijlage A", "Hoofdstkbijlagen");
            document.Paragraphs[0].OutlineLevel = 2;
        });

        var markdown = new WordDocumentService().ReadAsMarkdown(path);

        Assert.Contains("## Bijlage A", markdown);
    }

    [Fact]
    public void Outline_strips_the_rendered_heading_number_label()
    {
        // GetText renders the list label into the text — for an ODT the counters are not
        // resolved, so every level reads "1": section 1.2 comes out "1.1.Versiebeheer".
        var path = CreateNumberedHeadingDocument("Versiebeheer", listLevel: 1, "Heading2");

        var nodes = new WordDocumentService().GetOutline(path);

        Assert.Equal("Versiebeheer", Assert.Single(nodes).Text);
    }

    [Fact]
    public void Outline_keeps_a_typed_number_when_the_paragraph_is_not_in_a_list()
    {
        var path = TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "1.1 Versiebeheer", "Heading2");
        });

        var nodes = new WordDocumentService().GetOutline(path);

        Assert.Equal("1.1 Versiebeheer", Assert.Single(nodes).Text);
    }

    [Fact]
    public void Outline_keeps_a_number_the_author_typed_into_a_numbered_heading()
    {
        // Level 0 renders one segment ("1."); the "1.2.3" the author typed stays.
        var path = CreateNumberedHeadingDocument("1.2.3 Norm", listLevel: 0, "Heading1");

        var nodes = new WordDocumentService().GetOutline(path);

        Assert.Equal("1.2.3 Norm", Assert.Single(nodes).Text);
    }

    [Fact]
    public void ReadStructured_strips_the_rendered_heading_number_label()
    {
        var path = CreateNumberedHeadingDocument("Datawarehouse", listLevel: 1, "Heading2");

        var structured = new WordDocumentService().ReadStructured(path);

        Assert.Contains(structured.Blocks, b => b is McpOffice.Models.HeadingBlock { Text: "Datawarehouse" });
    }

    [Fact]
    public void ReadAsMarkdown_strips_the_rendered_heading_number_label()
    {
        var path = CreateNumberedHeadingDocument("Datawarehouse", listLevel: 1, "Heading2");

        var markdown = new WordDocumentService().ReadAsMarkdown(path);

        Assert.Contains("## Datawarehouse", markdown);
        Assert.DoesNotContain("1.1.", markdown);
    }

    private static string CreateNumberedHeadingDocument(string text, int listLevel, string styleName)
    {
        return TestWordDocuments.Create((Document document) =>
        {
            var abstractList = document.AbstractNumberingLists.Add();
            abstractList.NumberingType = NumberingType.MultiLevel;
            var list = document.NumberingLists.Add(abstractList.Index);

            TestWordDocuments.AppendParagraph(document, text, styleName);

            var paragraph = document.Paragraphs[0];
            paragraph.ListIndex = list.Index;
            paragraph.ListLevel = listLevel;
        });
    }

    [Fact]
    public void Convert_rejects_unknown_output_format()
    {
        var path = CreateOdtDocument();
        var output = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.xyz");
        var service = new WordDocumentService();

        var ex = Assert.Throws<McpException>(() => service.Convert(path, output, format: null));
        Assert.Contains("unsupported_format", ex.Message);
    }

    [Theory]
    [InlineData("C:\\a\\b.docx", "OpenXml")]
    [InlineData("C:\\a\\b.DOCX", "OpenXml")]
    [InlineData("C:\\a\\b.docm", "Docm")]
    [InlineData("C:\\a\\b.odt", "Odt")]
    [InlineData("C:\\a\\b.rtf", "Rtf")]
    [InlineData("C:\\a\\b.doc", "Doc")]
    [InlineData("C:\\a\\b.txt", "PlainText")]
    [InlineData("C:\\a\\b.unknown", "OpenXml")]
    [InlineData("C:\\a\\b", "OpenXml")]
    public void WordFormats_maps_extension_to_document_format(string path, string expected)
    {
        var format = WordFormats.ForPath(path);

        var expectedFormat = expected switch
        {
            "OpenXml" => RichEditFormat.OpenXml,
            "Docm" => RichEditFormat.Docm,
            "Odt" => RichEditFormat.Odt,
            "Rtf" => RichEditFormat.Rtf,
            "Doc" => RichEditFormat.Doc,
            "PlainText" => RichEditFormat.PlainText,
            _ => throw new ArgumentOutOfRangeException(nameof(expected), expected, null)
        };
        Assert.Equal(expectedFormat, format);
    }

    private static string CreateOdtDocument()
    {
        return TestWordDocuments.CreateOdt(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Introduction", "Heading 1");
            TestWordDocuments.AppendParagraph(document, "Background", "Heading 2");
            TestWordDocuments.AppendParagraph(document, "Plain paragraph");
        });
    }
}
