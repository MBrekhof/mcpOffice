using McpOffice.Models;
using McpOffice.Services.Pdf;

namespace McpOffice.Tests.Pdf;

public class LayoutTextRendererTests
{
    /// <summary>Monospaced 6pt-wide glyphs, so a word's x maps cleanly onto a character column.</summary>
    private static PdfWordBox Word(string text, double column, double y, double charWidth = 6)
        => new(text, column * charWidth, y, text.Length * charWidth, 10, null, null);

    private static PdfTextLine Line(double y, params PdfWordBox[] words)
        => new(1, words.Min(w => w.X), y, string.Join(' ', words.Select(w => w.Text)), words);

    [Fact]
    public void Empty_input_renders_empty()
        => Assert.Equal(string.Empty, LayoutTextRenderer.Render([]));

    [Fact]
    public void Words_are_padded_to_their_column()
    {
        var text = LayoutTextRenderer.Render([Line(100, Word("Parameter", 0, 100), Word("Eenheid", 20, 100))]);

        Assert.Equal("Parameter           Eenheid", text);
    }

    [Fact]
    public void Columns_line_up_across_rows()
    {
        var text = LayoutTextRenderer.Render(
        [
            Line(100, Word("Chloride", 0, 100), Word("mg/l", 20, 100)),
            Line(112, Word("Sulfaat", 0, 112), Word("mg/l", 20, 112)),
        ]);

        var rows = text.Split(Environment.NewLine);
        Assert.Equal(2, rows.Length);
        Assert.Equal(rows[0].IndexOf("mg/l", StringComparison.Ordinal),
                     rows[1].IndexOf("mg/l", StringComparison.Ordinal));
    }

    [Fact]
    public void Left_margin_is_removed_so_output_is_not_indented()
    {
        var text = LayoutTextRenderer.Render(
        [
            Line(100, Word("a", 10, 100), Word("b", 20, 100)),
            Line(112, Word("c", 10, 112)),
        ]);

        Assert.All(text.Split(Environment.NewLine), row => Assert.False(row.StartsWith(' ')));
    }

    [Fact]
    public void Vertical_gap_becomes_a_blank_line()
    {
        var text = LayoutTextRenderer.Render(
        [
            Line(100, Word("header", 0, 100)),
            Line(160, Word("body", 0, 160)),   // 6x line height below
        ]);

        var rows = text.Split(Environment.NewLine);
        Assert.Equal("header", rows[0]);
        Assert.Equal("body", rows[^1]);
        Assert.Contains(rows, string.IsNullOrEmpty);
    }

    [Fact]
    public void Adjacent_lines_get_no_blank_line()
    {
        var text = LayoutTextRenderer.Render(
        [
            Line(100, Word("one", 0, 100)),
            Line(112, Word("two", 0, 112)),
        ]);

        Assert.Equal($"one{Environment.NewLine}two", text);
    }

    [Fact]
    public void Overlapping_words_are_separated_rather_than_lost()
    {
        // Two words claiming the same column must both survive.
        var text = LayoutTextRenderer.Render([Line(100, Word("aaaaaa", 0, 100), Word("bbb", 1, 100))]);

        Assert.Contains("aaaaaa", text);
        Assert.Contains("bbb", text);
    }

    [Fact]
    public void Zero_width_words_fall_back_to_plain_text()
    {
        var word = new PdfWordBox("x", 0, 100, 0, 10, null, null);
        var text = LayoutTextRenderer.Render([new PdfTextLine(1, 0, 100, "x", [word])]);

        Assert.Equal("x", text);
    }
}
