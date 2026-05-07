using System.Text;
using McpOffice.Services.Excel.Csv;

namespace McpOffice.Tests.Excel.Csv;

public class CsvWriterTests
{
    [Fact]
    public void Writes_plain_text_rows_with_crlf_between_rows_and_no_trailing_crlf()
    {
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { "a", "b", "c" });
            writer.WriteRow(new object?[] { "d", "e", "f" });
        });

        var text = Encoding.UTF8.GetString(bytes);
        Assert.Equal("a,b,c\r\nd,e,f", text);
    }

    [Fact]
    public void Quotes_text_containing_comma_quote_or_newline()
    {
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { "no special", "has,comma", "has\"quote", "has\nnewline", "has\rcr" });
        });

        var text = Encoding.UTF8.GetString(bytes);
        Assert.Equal(
            "no special,\"has,comma\",\"has\"\"quote\",\"has\nnewline\",\"has\rcr\"",
            text);
    }

    [Fact]
    public void Numbers_use_invariant_culture_with_no_thousand_separators()
    {
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { 0.21, 1234567.89, 42, -7.5, 0.0 });
        });

        Assert.Equal("0.21,1234567.89,42,-7.5,0", Encoding.UTF8.GetString(bytes));
    }

    [Fact]
    public void DateTime_uses_iso_8601_with_T_separator_and_seconds()
    {
        var dt = new DateTime(2026, 5, 7, 14, 30, 0, DateTimeKind.Unspecified);
        var midnight = new DateTime(2026, 5, 7, 0, 0, 0, DateTimeKind.Unspecified);

        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { dt, midnight });
        });

        Assert.Equal("2026-05-07T14:30:00,2026-05-07T00:00:00", Encoding.UTF8.GetString(bytes));
    }

    [Fact]
    public void Boolean_emits_lowercase()
    {
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { true, false });
        });

        Assert.Equal("true,false", Encoding.UTF8.GetString(bytes));
    }

    [Fact]
    public void Null_and_empty_string_both_emit_empty_unquoted_field()
    {
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { null, "", "x" });
        });

        Assert.Equal(",,x", Encoding.UTF8.GetString(bytes));
    }

    [Fact]
    public void NlNL_host_culture_does_not_leak_into_output()
    {
        var original = System.Globalization.CultureInfo.CurrentCulture;
        try
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("nl-NL");
            var bytes = WriteToBytes(writer =>
            {
                writer.WriteRow(new object?[] { 0.21, 1234.5 });
            });
            Assert.Equal("0.21,1234.5", Encoding.UTF8.GetString(bytes));
        }
        finally
        {
            System.Globalization.CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void Output_is_utf8_without_bom()
    {
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { "héllo" });
        });

        // No UTF-8 BOM (EF BB BF) at the start.
        Assert.NotEqual(0xEF, bytes[0]);

        // Round-trips as UTF-8.
        Assert.Equal("héllo", Encoding.UTF8.GetString(bytes));
        Assert.Equal(new byte[] { 0x68, 0xC3, 0xA9, 0x6C, 0x6C, 0x6F }, bytes);
    }

    private static byte[] WriteToBytes(Action<CsvWriter> write)
    {
        using var stream = new MemoryStream();
        using (var writer = new CsvWriter(stream))
        {
            write(writer);
        }
        return stream.ToArray();
    }
}
