using McpOffice.Services.Pdf;
using ModelContextProtocol;

namespace McpOffice.Tests.Pdf;

public class PageRangeTests
{
    [Fact]
    public void Null_or_blank_selects_every_page()
    {
        Assert.Equal(new[] { 1, 2, 3 }, PageRange.Parse(null, 3));
        Assert.Equal(new[] { 1, 2, 3 }, PageRange.Parse("", 3));
        Assert.Equal(new[] { 1, 2, 3 }, PageRange.Parse("   ", 3));
    }

    [Fact]
    public void Single_page()
        => Assert.Equal(new[] { 2 }, PageRange.Parse("2", 5));

    [Fact]
    public void Inclusive_span()
        => Assert.Equal(new[] { 2, 3, 4 }, PageRange.Parse("2-4", 5));

    [Fact]
    public void Open_ended_span_runs_to_the_last_page()
        => Assert.Equal(new[] { 3, 4, 5 }, PageRange.Parse("3-", 5));

    [Fact]
    public void Mixed_list_is_sorted_and_deduplicated()
        => Assert.Equal(new[] { 1, 3, 4, 5, 9 }, PageRange.Parse("9,1,3-5,4", 10));

    [Fact]
    public void Whitespace_is_tolerated()
        => Assert.Equal(new[] { 1, 2, 7 }, PageRange.Parse(" 1 , 2 , 7 ", 8));

    [Fact]
    public void Empty_document_yields_nothing()
        => Assert.Empty(PageRange.Parse("1-3", 0));

    [Theory]
    [InlineData("0")]
    [InlineData("6")]
    [InlineData("1-6")]
    public void Page_outside_the_document_is_page_not_found(string range)
    {
        var ex = Assert.Throws<McpException>(() => PageRange.Parse(range, 5));
        Assert.StartsWith("[page_not_found]", ex.Message);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("1-x")]
    [InlineData("-3")]
    [InlineData("4-2")]
    public void Malformed_range_is_invalid_page_range(string range)
    {
        var ex = Assert.Throws<McpException>(() => PageRange.Parse(range, 5));
        Assert.StartsWith("[invalid_page_range]", ex.Message);
    }
}
