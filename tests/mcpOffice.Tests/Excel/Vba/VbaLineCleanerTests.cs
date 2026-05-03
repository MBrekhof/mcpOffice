using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaLineCleanerTests
{
    [Fact]
    public void Strips_apostrophe_comment()
    {
        var lines = VbaLineCleaner.Clean("x = 1 ' set x");
        Assert.Single(lines);
        Assert.Equal("x = 1", lines[0].Text.TrimEnd());
        Assert.Equal(1, lines[0].LineNumber);
    }

    [Fact]
    public void Apostrophe_inside_string_is_not_a_comment()
    {
        var lines = VbaLineCleaner.Clean("s = \"isn't a comment\"");
        Assert.Single(lines);
        Assert.Contains("<STR>", lines[0].Text);
        Assert.DoesNotContain("isn't", lines[0].Text);
    }

    [Fact]
    public void Doubled_quote_escape_inside_string()
    {
        var lines = VbaLineCleaner.Clean("s = \"he said \"\"hi\"\"\"");
        Assert.Single(lines);
        Assert.Contains("<STR>", lines[0].Text);
        Assert.DoesNotContain("he said", lines[0].Text);
    }

    [Fact]
    public void Rem_statement_is_treated_as_comment()
    {
        var lines = VbaLineCleaner.Clean("Rem this is a comment");
        Assert.Single(lines);
        Assert.Equal("", lines[0].Text.Trim());
    }

    [Fact]
    public void Folds_underscore_continuation()
    {
        var src = "Sub Foo(x As Long, _\r\n            y As Long)";
        var lines = VbaLineCleaner.Clean(src);
        Assert.Single(lines);
        Assert.Contains("Sub Foo(x As Long,", lines[0].Text);
        Assert.Contains("y As Long)", lines[0].Text);
        Assert.Equal(1, lines[0].LineNumber);
    }

    [Fact]
    public void Preserves_originalText_for_string_literal_capture()
    {
        var lines = VbaLineCleaner.Clean("Set ws = Worksheets(\"Data\")");
        Assert.Single(lines);
        Assert.Contains("\"Data\"", lines[0].OriginalText);
        Assert.Contains("<STR>", lines[0].Text);
    }
}
