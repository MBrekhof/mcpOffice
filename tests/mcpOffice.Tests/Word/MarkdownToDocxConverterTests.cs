using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class MarkdownToDocxConverterTests
{
    [Fact]
    public void Apply_with_empty_markdown_does_not_throw()
    {
        using var server = new RichEditDocumentServer();
        // Should not throw
        MarkdownToDocxConverter.Apply(server.Document, string.Empty, null);
    }

    [Fact]
    public void Empty_markdown_produces_empty_document()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "", null);
        Assert.Single(server.Document.Paragraphs);
        Assert.Equal(string.Empty, server.Document.GetText(server.Document.Range).Trim());
    }

    [Fact]
    public void Plain_paragraphs_become_paragraphs_with_literal_text()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "hello world\n\nsecond para", null);
        var text = server.Document.GetText(server.Document.Range).Trim();
        Assert.Contains("hello world", text);
        Assert.Contains("second para", text);
        Assert.True(server.Document.Paragraphs.Count >= 2,
            $"expected ≥2 paragraphs, got {server.Document.Paragraphs.Count}");
    }

    [Fact]
    public void Headings_1_through_6_get_correct_paragraph_style()
    {
        var md = "# h1\n\n## h2\n\n### h3\n\n#### h4\n\n##### h5\n\n###### h6";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        var headingParas = server.Document.Paragraphs
            .Where(p => p.Style?.Name?.StartsWith("Heading ") == true)
            .ToList();

        Assert.Equal(6, headingParas.Count);
        for (int i = 0; i < 6; i++)
        {
            Assert.Equal($"Heading {i + 1}", headingParas[i].Style!.Name);
        }
    }

    [Fact]
    public void Unordered_list_produces_bulleted_paragraphs()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "- a\n- b\n- c", null);
        var listParas = server.Document.Paragraphs
            .Where(p => p.ListIndex >= 0)
            .ToList();
        Assert.Equal(3, listParas.Count);
        Assert.All(listParas, p => Assert.Equal(0, p.ListLevel));
    }

    [Fact]
    public void Ordered_list_produces_numbered_paragraphs()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "1. one\n2. two", null);
        var listParas = server.Document.Paragraphs
            .Where(p => p.ListIndex >= 0)
            .ToList();
        Assert.Equal(2, listParas.Count);
    }

    [Fact]
    public void Nested_list_indents_per_depth()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "- outer\n  - inner", null);
        var paras = server.Document.Paragraphs.Where(p => p.ListIndex >= 0).ToList();
        Assert.Equal(2, paras.Count);
        Assert.Equal(0, paras[0].ListLevel);
        Assert.Equal(1, paras[1].ListLevel);
    }

    [Fact]
    public void Blockquote_indents_left_quarter_inch()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "> quoted text", null);
        var doc = server.Document;
        var quotedPara = doc.Paragraphs
            .FirstOrDefault(p => doc.GetText(p.Range).Contains("quoted text"));
        Assert.NotNull(quotedPara);
        Assert.True(quotedPara!.LeftIndent > 0,
            $"expected LeftIndent > 0, got {quotedPara.LeftIndent}");
    }

    [Fact]
    public void Fenced_code_block_each_line_is_monospace_paragraph()
    {
        var md = "```\nfoo\nbar\n```";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);
        var doc = server.Document;

        var codeParas = doc.Paragraphs
            .Where(p => doc.GetText(p.Range).Trim() is "foo" or "bar")
            .ToList();
        Assert.Equal(2, codeParas.Count);

        foreach (var p in codeParas)
        {
            var content = doc.GetText(p.Range);
            var firstNonWs = 0;
            while (firstNonWs < content.Length && char.IsWhiteSpace(content[firstNonWs])) firstNonWs++;
            if (firstNonWs >= content.Length) continue;
            var pos = doc.CreatePosition(p.Range.Start.ToInt() + firstNonWs);
            var range = doc.CreateRange(pos, 1);
            var props = doc.BeginUpdateCharacters(range);
            try { Assert.Equal("Consolas", props.FontName); }
            finally { doc.EndUpdateCharacters(props); }
        }
    }

    [Fact]
    public void Fenced_code_block_background_is_light_grey()
    {
        var md = "```\nhello\n```";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);
        var doc = server.Document;

        var codePara = doc.Paragraphs
            .FirstOrDefault(p => doc.GetText(p.Range).Trim() == "hello");
        Assert.NotNull(codePara);

        var content = doc.GetText(codePara!.Range);
        var firstNonWs = 0;
        while (firstNonWs < content.Length && char.IsWhiteSpace(content[firstNonWs])) firstNonWs++;
        var pos = doc.CreatePosition(codePara.Range.Start.ToInt() + firstNonWs);
        var range = doc.CreateRange(pos, 1);
        var props = doc.BeginUpdateCharacters(range);
        try
        {
            var bg = props.BackColor;
            Assert.True(bg.HasValue, "expected BackColor to be set");
            var c = bg!.Value;
            Assert.True(c.R == 0xF2 && c.G == 0xF2 && c.B == 0xF2,
                $"expected #F2F2F2 background, got R={c.R} G={c.G} B={c.B}");
        }
        finally { doc.EndUpdateCharacters(props); }
    }

    [Fact]
    public void Hr_emits_paragraph_with_bottom_border()
    {
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, "before\n\n---\n\nafter", null);
        var doc = server.Document;
        var hrPara = doc.Paragraphs
            .FirstOrDefault(p =>
            {
                var props = doc.BeginUpdateParagraphs(p.Range);
                try
                {
                    return props.Borders.BottomBorder.LineStyle != BorderLineStyle.None;
                }
                finally { doc.EndUpdateParagraphs(props); }
            });
        Assert.NotNull(hrPara);
    }

    [Fact]
    public void Pipe_table_creates_real_table_with_bold_shaded_header()
    {
        var md = "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);
        var doc = server.Document;

        Assert.Single(doc.Tables);
        var table = doc.Tables[0];
        Assert.Equal(3, table.Rows.Count);                // header + 2 data rows
        Assert.Equal(2, table.Rows[0].Cells.Count);

        // Header cell shading == #F2F2F2
        var headerCell = table.Rows[0].Cells[0];
        var bg = headerCell.BackgroundColor;
        Assert.Equal(System.Drawing.Color.FromArgb(0xF2, 0xF2, 0xF2).ToArgb(), bg.ToArgb());

        // Header runs bold — check via BeginUpdateCharacters on the content range
        var props = doc.BeginUpdateCharacters(headerCell.ContentRange);
        try { Assert.Equal(true, props.Bold); }
        finally { doc.EndUpdateCharacters(props); }
    }

    [Fact]
    public void Pipe_table_column_alignment_from_gfm_spec()
    {
        var md = "| L | C | R |\n|:---|:---:|---:|\n| a | b | c |";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        var doc = server.Document;
        var table = doc.Tables[0];
        var dataRow = table.Rows[1];

        Assert.Equal(ParagraphAlignment.Left,   AlignmentOf(doc, dataRow.Cells[0]));
        Assert.Equal(ParagraphAlignment.Center, AlignmentOf(doc, dataRow.Cells[1]));
        Assert.Equal(ParagraphAlignment.Right,  AlignmentOf(doc, dataRow.Cells[2]));

        static ParagraphAlignment AlignmentOf(Document doc, TableCell cell)
        {
            var props = doc.BeginUpdateParagraphs(cell.ContentRange);
            try { return props.Alignment ?? ParagraphAlignment.Left; }
            finally { doc.EndUpdateParagraphs(props); }
        }
    }

    [Fact]
    public void Indented_code_block_each_line_is_monospace_paragraph()
    {
        // Four-space indent = code block in Markdown
        var md = "    alpha\n    beta";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);
        var doc = server.Document;

        var codeParas = doc.Paragraphs
            .Where(p => doc.GetText(p.Range).Trim() is "alpha" or "beta")
            .ToList();
        Assert.Equal(2, codeParas.Count);

        foreach (var p in codeParas)
        {
            var content = doc.GetText(p.Range);
            var firstNonWs = 0;
            while (firstNonWs < content.Length && char.IsWhiteSpace(content[firstNonWs])) firstNonWs++;
            if (firstNonWs >= content.Length) continue;
            var pos = doc.CreatePosition(p.Range.Start.ToInt() + firstNonWs);
            var range = doc.CreateRange(pos, 1);
            var props = doc.BeginUpdateCharacters(range);
            try { Assert.Equal("Consolas", props.FontName); }
            finally { doc.EndUpdateCharacters(props); }
        }
    }
}
