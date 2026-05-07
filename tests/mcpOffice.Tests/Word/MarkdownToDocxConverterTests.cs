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
    public void Table_cells_render_inline_formatting()
    {
        var md = "| Name | Status |\n|---|---|\n| `Foo()` | **active** |";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);
        var doc = server.Document;

        Assert.Single(doc.Tables);
        var table = doc.Tables[0];

        // Data row, first cell: "Foo()" should be in Consolas (came from `Foo()` inline code)
        var nameCell = table.Rows[1].Cells[0];
        var nameText = doc.GetText(nameCell.ContentRange);
        var fooIdx = nameText.IndexOf("Foo()", StringComparison.Ordinal);
        Assert.True(fooIdx >= 0, $"expected 'Foo()' in cell text, got: '{nameText}'");
        var fooRange = doc.CreateRange(nameCell.ContentRange.Start.ToInt() + fooIdx, "Foo()".Length);
        var nameProps = doc.BeginUpdateCharacters(fooRange);
        try { Assert.Equal("Consolas", nameProps.FontName); }
        finally { doc.EndUpdateCharacters(nameProps); }

        // Data row, second cell: "active" should be bold (from **active**)
        var statusCell = table.Rows[1].Cells[1];
        var statusText = doc.GetText(statusCell.ContentRange);
        var activeIdx = statusText.IndexOf("active", StringComparison.Ordinal);
        Assert.True(activeIdx >= 0, $"expected 'active' in cell text, got: '{statusText}'");
        var activeRange = doc.CreateRange(statusCell.ContentRange.Start.ToInt() + activeIdx, "active".Length);
        var statusProps = doc.BeginUpdateCharacters(activeRange);
        try { Assert.True(statusProps.Bold == true, $"expected Bold=true, got {statusProps.Bold}"); }
        finally { doc.EndUpdateCharacters(statusProps); }
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
    public void Emphasis_produces_bold_italic_runs()
    {
        var md = "**bold** *italic* ***both***";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);
        var doc = server.Document;

        Assert.Contains("bold", doc.GetText(doc.Range));
        Assert.Contains("italic", doc.GetText(doc.Range));
        Assert.Contains("both", doc.GetText(doc.Range));

        Assert.True(HasRunMatching(doc, p => p.Bold == true && p.Italic != true),
            "expected at least one bold-only run");
        Assert.True(HasRunMatching(doc, p => p.Italic == true && p.Bold != true),
            "expected at least one italic-only run");
        Assert.True(HasRunMatching(doc, p => p.Bold == true && p.Italic == true),
            "expected at least one bold-italic run");

        static bool HasRunMatching(Document doc, Func<CharacterProperties, bool> pred)
        {
            var totalChars = doc.Range.End.ToInt() - doc.Range.Start.ToInt();
            for (int i = 0; i < totalChars; i++)
            {
                var range = doc.CreateRange(doc.Range.Start.ToInt() + i, 1);
                var props = doc.BeginUpdateCharacters(range);
                try { if (pred(props)) return true; }
                finally { doc.EndUpdateCharacters(props); }
            }
            return false;
        }
    }

    [Fact]
    public void Inline_code_run_uses_Consolas()
    {
        var md = "x `code` y";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        var doc = server.Document;
        var fullText = doc.GetText(doc.Range);
        var idx = fullText.IndexOf("code", StringComparison.Ordinal);
        Assert.True(idx >= 0, $"expected 'code' in document text, got: {fullText}");

        var range = doc.CreateRange(doc.Range.Start.ToInt() + idx, "code".Length);
        var props = doc.BeginUpdateCharacters(range);
        try
        {
            Assert.Equal("Consolas", props.FontName);
        }
        finally { doc.EndUpdateCharacters(props); }
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

    [Fact]
    public void Hyperlink_emits_field_with_target()
    {
        var md = "see [the docs](https://example.com/x) here";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        Assert.True(server.Document.Hyperlinks.Count > 0,
            $"expected at least one hyperlink, got {server.Document.Hyperlinks.Count}");
        Assert.Equal("https://example.com/x", server.Document.Hyperlinks[0].NavigateUri);
    }

    [Fact]
    public void Autolink_emits_hyperlink_with_url_as_text()
    {
        var md = "see <https://example.com/y>";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);
        Assert.Contains(server.Document.Hyperlinks.Cast<Hyperlink>(),
            h => h.NavigateUri == "https://example.com/y");
    }

    [Fact]
    public void Hard_break_inserts_line_break_inside_paragraph()
    {
        // Two trailing spaces + newline = hard break in CommonMark.
        // Both pieces of text must appear inside a SINGLE paragraph (not split across two paragraphs),
        // because a hard break is a line-break-within-paragraph (\v in DevExpress internal model).
        // Note: Document.GetText normalises \v to \r\n on output, so we verify structure rather than raw char.
        var md = "line one  \nline two";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        var doc = server.Document;
        var text = doc.GetText(doc.Range);
        Assert.Contains("line one", text);
        Assert.Contains("line two", text);

        // Both words must live in the SAME paragraph (hard break, not a new paragraph).
        var paraWithBoth = doc.Paragraphs
            .FirstOrDefault(p =>
            {
                var t = doc.GetText(p.Range);
                return t.Contains("line one") && t.Contains("line two");
            });
        Assert.NotNull(paraWithBoth);
    }

    [Fact]
    public void Soft_break_becomes_a_space()
    {
        // Single newline (no trailing spaces) = soft break in CommonMark.
        // Both pieces should appear in the same paragraph, separated by a space.
        var md = "line one\nline two";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        var doc = server.Document;
        var text = doc.GetText(doc.Range);
        // Both words must be present.
        Assert.Contains("line one", text);
        Assert.Contains("line two", text);
        // Should be only one real (non-empty) paragraph.
        var nonEmpty = doc.Paragraphs
            .Where(p => doc.GetText(p.Range).Trim().Length > 0)
            .ToList();
        Assert.Single(nonEmpty);
    }

    [Fact]
    public void Image_local_file_is_embedded()
    {
        using var tmpDir = new TempDir();
        var pngPath = Path.Combine(tmpDir.Path, "dot.png");
        File.WriteAllBytes(pngPath, OnePixelPng());
        var md = "![dot](dot.png)";

        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, tmpDir.Path);
        Assert.Single(server.Document.Images);
    }

    [Fact]
    public void Image_missing_local_file_is_dropped_no_throw()
    {
        using var tmpDir = new TempDir();
        var md = "![](missing.png)";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, tmpDir.Path);
        Assert.Empty(server.Document.Images);
    }

    [Fact]
    public void Image_remote_url_is_dropped_no_throw()
    {
        var md = "![](https://example.com/x.png)";
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);
        Assert.Empty(server.Document.Images);
    }

    private static byte[] OnePixelPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    private sealed class TempDir : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            $"mcpoffice-md-{Guid.NewGuid():N}");
        public TempDir() => Directory.CreateDirectory(Path);
        public void Dispose() { try { Directory.Delete(Path, true); } catch { } }
    }
}
