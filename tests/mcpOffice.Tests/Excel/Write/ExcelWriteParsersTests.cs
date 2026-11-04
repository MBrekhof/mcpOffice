using System.Drawing;
using McpOffice.Services.Excel.Write;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Write;

public class ExcelWriteParsersTests
{
    [Theory]
    [InlineData("A", 1)]
    [InlineData("c", 3)]
    [InlineData("AA", 27)]
    [InlineData("3", 3)]
    [InlineData("XFD", 16384)]
    public void ParseColumn_accepts_letters_and_numbers(string input, int expected) =>
        Assert.Equal(expected, ExcelWriteParsers.ParseColumn(input));

    [Theory]
    [InlineData("0")]
    [InlineData("XFE")]
    [InlineData("16385")]
    [InlineData("")]
    [InlineData("1A")]
    [InlineData("A-1")]
    public void ParseColumn_rejects(string input) =>
        Assert.Throws<WriteValidationException>(() => ExcelWriteParsers.ParseColumn(input));

    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(16384, "XFD")]
    public void ColumnLetter_round_trips(int column, string expected)
    {
        Assert.Equal(expected, ExcelWriteParsers.ColumnLetter(column));
        Assert.Equal(column, ExcelWriteParsers.ParseColumn(expected));
    }

    [Theory]
    [InlineData("B2", 2, 2)]
    [InlineData("a1", 1, 1)]
    [InlineData("$C$10", 10, 3)]
    [InlineData("XFD1048576", 1048576, 16384)]
    public void ParseCell_is_one_based(string a1, int row, int column) =>
        Assert.Equal((row, column), ExcelWriteParsers.ParseCell(a1));

    [Theory]
    [InlineData("A0")]
    [InlineData("1A")]
    [InlineData("")]
    [InlineData("A1:B2")]
    [InlineData("A1048577")]
    [InlineData("XFE1")]
    public void ParseCell_rejects(string a1) =>
        Assert.Throws<WriteValidationException>(() => ExcelWriteParsers.ParseCell(a1));

    [Theory]
    [InlineData("A", 1, 1)]
    [InlineData("A:C", 1, 3)]
    [InlineData("2:4", 2, 4)]
    public void ParseColumnSpan(string input, int first, int last) =>
        Assert.Equal((first, last), ExcelWriteParsers.ParseColumnSpan(input));

    [Theory]
    [InlineData("C:A")]
    [InlineData("A:B:C")]
    [InlineData("")]
    public void ParseColumnSpan_rejects(string input) =>
        Assert.Throws<WriteValidationException>(() => ExcelWriteParsers.ParseColumnSpan(input));

    [Theory]
    [InlineData("Data")]
    [InlineData("Résultats 2026")]
    [InlineData("a234567890123456789012345678901")]   // 31 chars
    public void ValidateSheetName_accepts(string name) =>
        ExcelWriteParsers.ValidateSheetName(name);

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("a2345678901234567890123456789012")]   // 32 chars
    [InlineData("a/b")]
    [InlineData("a:b")]
    [InlineData("a?b")]
    [InlineData("a*b")]
    [InlineData("a[b]")]
    [InlineData("'x")]
    [InlineData("x'")]
    public void ValidateSheetName_rejects(string name) =>
        Assert.Throws<WriteValidationException>(() => ExcelWriteParsers.ValidateSheetName(name));

    [Theory]
    [InlineData("#D9D9D9", 0xD9, 0xD9, 0xD9)]
    [InlineData("#ff0000", 0xFF, 0, 0)]
    public void ParseHexColor_accepts(string hex, int r, int g, int b) =>
        Assert.Equal(Color.FromArgb(r, g, b), ExcelWriteParsers.ParseHexColor(hex));

    [Theory]
    [InlineData("red")]
    [InlineData("#FFF")]
    [InlineData("FF0000")]
    [InlineData("#GG0000")]
    [InlineData(null)]
    public void ParseHexColor_rejects(string? hex) =>
        Assert.Throws<WriteValidationException>(() => ExcelWriteParsers.ParseHexColor(hex));

    [Fact]
    public void ToolError_InvalidOperation_prefix()
    {
        var ex = Assert.IsType<McpException>(ToolError.InvalidOperation("operations[3] (formatRange)", "fillColor 'red' is not #RRGGBB"));
        Assert.StartsWith("[invalid_operation] operations[3] (formatRange): ", ex.Message);
        Assert.EndsWith("fillColor 'red' is not #RRGGBB", ex.Message);
    }
}
