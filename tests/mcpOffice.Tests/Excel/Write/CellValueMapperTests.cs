using System.Text.Json;
using McpOffice.Services.Excel.Write;

namespace McpOffice.Tests.Excel.Write;

public class CellValueMapperTests
{
    private static MappedCell Map(string json) => CellValueMapper.Map(JsonDocument.Parse(json).RootElement);

    [Fact]
    public void Number_boolean_null_by_json_type()
    {
        Assert.Equal(new MappedCell(MappedCellKind.Number, 12.5d, null), Map("12.5"));
        Assert.Equal(new MappedCell(MappedCellKind.Number, 40d, null), Map("40"));
        Assert.Equal(new MappedCell(MappedCellKind.Boolean, true, null), Map("true"));
        Assert.Equal(new MappedCell(MappedCellKind.Boolean, false, null), Map("false"));
        Assert.Equal(new MappedCell(MappedCellKind.Empty, null, null), Map("null"));
    }

    [Fact]
    public void Formula_is_any_string_starting_with_equals() =>
        Assert.Equal(new MappedCell(MappedCellKind.Formula, "=SUM(A1:A2)", null), Map("\"=SUM(A1:A2)\""));

    [Fact]
    public void Iso_date_becomes_date_with_default_format() =>
        Assert.Equal(new MappedCell(MappedCellKind.DateTime, new DateTime(2026, 11, 4), "yyyy-mm-dd"), Map("\"2026-11-04\""));

    [Theory]
    [InlineData("2026-11-04T10:15")]
    [InlineData("2026-11-04 10:15")]
    public void Iso_datetime_to_the_minute(string s) =>
        Assert.Equal(new MappedCell(MappedCellKind.DateTime, new DateTime(2026, 11, 4, 10, 15, 0), "yyyy-mm-dd hh:mm"), Map($"\"{s}\""));

    [Fact]
    public void Fractional_seconds_are_kept_in_the_value()
    {
        var cell = Map("\"2026-11-04T10:15:30.123\"");
        Assert.Equal(MappedCellKind.DateTime, cell.Kind);
        Assert.Equal(new DateTime(2026, 11, 4, 10, 15, 30, 123), cell.Value);
        Assert.Equal("yyyy-mm-dd hh:mm:ss", cell.DefaultNumberFormat);
    }

    [Theory]
    [InlineData("2026-11-04T10:15:30Z")]
    [InlineData("2026-11-04T10:15:30+01:00")]
    [InlineData("2026-11-04 10:15:30-05:00")]
    public void Zone_suffix_is_dropped_and_wall_clock_kept(string s) =>
        Assert.Equal(new MappedCell(MappedCellKind.DateTime, new DateTime(2026, 11, 4, 10, 15, 30), "yyyy-mm-dd hh:mm:ss"), Map($"\"{s}\""));

    [Theory]
    [InlineData("12.5")]
    [InlineData("2026-13-45")]
    [InlineData("2026-11-04T25:00")]
    [InlineData("hello")]
    [InlineData("04-11-2026")]
    [InlineData("")]
    public void Other_strings_stay_text(string s) =>
        Assert.Equal(new MappedCell(MappedCellKind.Text, s, null), Map(JsonSerializer.Serialize(s)));

    [Theory]
    [InlineData("'=not a formula", "=not a formula")]
    [InlineData("'2026-11-04", "2026-11-04")]
    [InlineData("''quoted", "'quoted")]
    public void Apostrophe_escapes_to_text(string input, string expected) =>
        Assert.Equal(new MappedCell(MappedCellKind.Text, expected, null), Map(JsonSerializer.Serialize(input)));

    [Theory]
    [InlineData("{\"a\":1}")]
    [InlineData("[1,2]")]
    public void Object_or_array_value_is_rejected(string json) =>
        Assert.Throws<WriteValidationException>(() => Map(json));

    [Fact]
    public void Text_over_32767_chars_is_rejected()
    {
        Assert.Equal(MappedCellKind.Text, Map(JsonSerializer.Serialize(new string('x', 32767))).Kind);
        Assert.Throws<WriteValidationException>(() => Map(JsonSerializer.Serialize(new string('x', 32768))));
    }
}
