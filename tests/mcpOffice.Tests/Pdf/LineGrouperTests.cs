using McpOffice.Models;
using McpOffice.Services.Pdf;

namespace McpOffice.Tests.Pdf;

public class LineGrouperTests
{
    private static PdfWordBox Word(string text, double x, double y, double height = 10, double width = 20)
        => new(text, x, y, width, height, null, null);

    [Fact]
    public void Empty_input_yields_no_lines()
        => Assert.Empty(LineGrouper.Group(1, []));

    [Fact]
    public void Words_at_the_same_y_form_one_line_ordered_left_to_right()
    {
        var lines = LineGrouper.Group(1,
        [
            Word("world", 120, 100),
            Word("hello", 40, 100),
        ]);

        var line = Assert.Single(lines);
        Assert.Equal("hello world", line.Text);
        Assert.Equal(40, line.X);
        Assert.Equal(1, line.PageNumber);
    }

    [Fact]
    public void Words_far_apart_vertically_split_into_lines_in_reading_order()
    {
        var lines = LineGrouper.Group(1,
        [
            Word("second", 40, 130),
            Word("first", 40, 100),
        ]);

        Assert.Equal(2, lines.Count);
        Assert.Equal("first", lines[0].Text);
        Assert.Equal("second", lines[1].Text);
    }

    [Fact]
    public void Small_vertical_jitter_stays_on_one_line()
    {
        // Real PDFs rarely place every glyph on exactly the same baseline.
        var lines = LineGrouper.Group(1,
        [
            Word("a", 10, 100),
            Word("b", 40, 101.5),
            Word("c", 70, 99),
        ]);

        var line = Assert.Single(lines);
        Assert.Equal("a b c", line.Text);
    }

    [Fact]
    public void Tolerance_scales_with_word_height()
    {
        // A 3pt offset splits 4pt-tall text but not 40pt-tall text.
        var small = LineGrouper.Group(1, [Word("a", 10, 100, height: 4), Word("b", 40, 103, height: 4)]);
        var large = LineGrouper.Group(1, [Word("a", 10, 100, height: 40), Word("b", 40, 103, height: 40)]);

        Assert.Equal(2, small.Count);
        Assert.Single(large);
    }

    [Fact]
    public void Line_y_is_the_topmost_word_and_words_are_preserved()
    {
        var lines = LineGrouper.Group(3, [Word("x", 10, 102), Word("y", 40, 100)]);

        var line = Assert.Single(lines);
        Assert.Equal(100, line.Y);
        Assert.Equal(3, line.PageNumber);
        Assert.Equal(2, line.Words.Count);
        Assert.Equal("x", line.Words[0].Text);
    }
}
