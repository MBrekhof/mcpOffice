using System.Globalization;
using System.Text;

namespace McpOffice.Services.Excel.Csv;

internal sealed class CsvWriter : IDisposable
{
    // UTF-8 without BOM. pandas.read_csv default; BOM breaks naive consumers.
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);
    private const string LineSeparator = "\r\n";

    private readonly StreamWriter _writer;
    private bool _firstRow = true;

    public CsvWriter(Stream stream)
    {
        _writer = new StreamWriter(stream, Utf8NoBom, bufferSize: 64 * 1024, leaveOpen: true);
    }

    public void WriteRow(IReadOnlyList<object?> values)
    {
        if (!_firstRow) _writer.Write(LineSeparator);
        _firstRow = false;

        for (var i = 0; i < values.Count; i++)
        {
            if (i > 0) _writer.Write(',');
            _writer.Write(FormatField(values[i]));
        }
    }

    private static string FormatField(object? value)
    {
        return value switch
        {
            null                 => string.Empty,
            bool b               => b ? "true" : "false",
            DateTime dt          => dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
            DateTimeOffset dto   => dto.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
            IFormattable f       => Quote(f.ToString(format: null, formatProvider: CultureInfo.InvariantCulture)),
            string s             => Quote(s),
            _                    => Quote(value.ToString() ?? string.Empty),
        };
    }

    private static string Quote(string text)
    {
        if (text.Length == 0) return text;
        if (text.IndexOfAny(['"', ',', '\r', '\n']) < 0) return text;
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }

    public void Dispose() => _writer.Dispose();
}
