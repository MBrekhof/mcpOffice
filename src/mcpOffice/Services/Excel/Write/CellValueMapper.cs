using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace McpOffice.Services.Excel.Write;

internal enum MappedCellKind { Empty, Number, Boolean, Text, Formula, DateTime }

internal sealed record MappedCell(MappedCellKind Kind, object? Value, string? DefaultNumberFormat);

/// <summary>
/// JSON value → cell, by JSON type only (design doc, "Cell values"). The only strings that become
/// something else are formulas (<c>=</c>), ISO dates and date-times; <c>'</c> escapes to text.
/// </summary>
internal static partial class CellValueMapper
{
    public const int MaxTextLength = 32767;   // Excel's cell limit
    public const int MaxFormulaLength = 8192; // Excel's formula limit

    private static readonly string[] DateFormats = ["yyyy-MM-dd"];
    private static readonly string[] DateTimeMinuteFormats = ["yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd HH:mm"];
    private static readonly string[] DateTimeSecondFormats =
    [
        "yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", "yyyy-MM-dd HH:mm:ss.FFFFFFF",
    ];

    // Excel has no time zones: a trailing Z or ±HH:mm is dropped and the wall-clock time kept (designer's call).
    [GeneratedRegex(@"(Z|[+-]\d{2}:\d{2})$")]
    private static partial Regex ZoneSuffixRegex();

    public static MappedCell Map(JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Undefined:
            case JsonValueKind.Null:
                return new MappedCell(MappedCellKind.Empty, null, null);
            case JsonValueKind.Number:
                return new MappedCell(MappedCellKind.Number, value.GetDouble(), null);
            case JsonValueKind.True:
                return new MappedCell(MappedCellKind.Boolean, true, null);
            case JsonValueKind.False:
                return new MappedCell(MappedCellKind.Boolean, false, null);
            case JsonValueKind.String:
                return MapString(value.GetString()!);
            default:
                throw new WriteValidationException($"a cell value must be a number, string, boolean or null, not a JSON {value.ValueKind.ToString().ToLowerInvariant()}");
        }
    }

    private static MappedCell MapString(string s)
    {
        if (s.StartsWith('\'')) return Text(s[1..]);
        if (s.StartsWith('='))
        {
            if (s.Length > MaxFormulaLength)
                throw new WriteValidationException($"formula is {s.Length} characters; Excel formulas hold at most {MaxFormulaLength}");
            return new MappedCell(MappedCellKind.Formula, s, null);
        }

        // Only strings shaped like a date are worth the parse; everything else is text.
        if (s.Length >= 10 && s.Length <= 40 && char.IsAsciiDigit(s[0]))
        {
            var withoutZone = ZoneSuffixRegex().Replace(s, "");
            var style = DateTimeStyles.None;
            if (DateTime.TryParseExact(withoutZone, DateFormats, CultureInfo.InvariantCulture, style, out var d))
                return new MappedCell(MappedCellKind.DateTime, d, "yyyy-mm-dd");
            if (DateTime.TryParseExact(withoutZone, DateTimeMinuteFormats, CultureInfo.InvariantCulture, style, out d))
                return new MappedCell(MappedCellKind.DateTime, d, "yyyy-mm-dd hh:mm");
            if (DateTime.TryParseExact(withoutZone, DateTimeSecondFormats, CultureInfo.InvariantCulture, style, out d))
                return new MappedCell(MappedCellKind.DateTime, d, "yyyy-mm-dd hh:mm:ss");
        }

        return Text(s);
    }

    private static MappedCell Text(string s)
    {
        if (s.Length > MaxTextLength)
            throw new WriteValidationException($"text is {s.Length} characters; Excel cells hold at most {MaxTextLength}");
        return new MappedCell(MappedCellKind.Text, s, null);
    }
}
