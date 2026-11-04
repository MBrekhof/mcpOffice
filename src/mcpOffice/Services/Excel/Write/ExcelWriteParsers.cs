using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;

namespace McpOffice.Services.Excel.Write;

/// <summary>Pure A1 / name / colour parsing for the write tools. Everything is 1-based; bad input throws <see cref="WriteValidationException"/>.</summary>
internal static partial class ExcelWriteParsers
{
    public const int MaxColumns = 16384;
    public const int MaxRows = 1_048_576;

    private static readonly char[] ForbiddenSheetNameChars = [':', '\\', '/', '?', '*', '[', ']'];

    [GeneratedRegex(@"^\$?(?<col>[A-Za-z]{1,3})\$?(?<row>[0-9]{1,7})$")]
    private static partial Regex CellRegex();

    [GeneratedRegex(@"^#[0-9A-Fa-f]{6}$")]
    private static partial Regex HexColorRegex();

    /// <summary>"A" → 1, "c" → 3, "AA" → 27, "3" → 3.</summary>
    public static int ParseColumn(string letterOrNumber)
    {
        var s = (letterOrNumber ?? "").Trim();
        if (s.Length == 0) throw new WriteValidationException("column is empty");

        int column;
        if (int.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out var number))
        {
            column = number;
        }
        else
        {
            if (s.Length > 3 || !s.All(char.IsAsciiLetter))
                throw new WriteValidationException($"column '{letterOrNumber}' is not a letter (A..XFD) or a number (1..{MaxColumns})");
            column = 0;
            foreach (var c in s) column = column * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }

        if (column < 1 || column > MaxColumns)
            throw new WriteValidationException($"column '{letterOrNumber}' is outside 1..{MaxColumns} (A..XFD)");
        return column;
    }

    public static string ColumnLetter(int column)
    {
        if (column < 1 || column > MaxColumns) throw new WriteValidationException($"column {column} is outside 1..{MaxColumns}");
        var s = "";
        while (column > 0)
        {
            column--;
            s = (char)('A' + column % 26) + s;
            column /= 26;
        }
        return s;
    }

    /// <summary>"B2" → (2, 2).</summary>
    public static (int Row, int Column) ParseCell(string a1)
    {
        var m = CellRegex().Match((a1 ?? "").Trim());
        if (!m.Success) throw new WriteValidationException($"'{a1}' is not a cell reference like B2");
        var row = int.Parse(m.Groups["row"].Value, CultureInfo.InvariantCulture);
        if (row < 1 || row > MaxRows) throw new WriteValidationException($"row in '{a1}' is outside 1..{MaxRows}");
        return (row, ParseColumn(m.Groups["col"].Value));
    }

    /// <summary>"A" → (1, 1); "A:C" → (1, 3).</summary>
    public static (int First, int Last) ParseColumnSpan(string columns)
    {
        var parts = (columns ?? "").Split(':');
        if (parts.Length > 2) throw new WriteValidationException($"'{columns}' is not a column or a column span like A:C");
        var first = ParseColumn(parts[0]);
        var last = parts.Length == 2 ? ParseColumn(parts[1]) : first;
        if (last < first) throw new WriteValidationException($"column span '{columns}' ends before it starts");
        return (first, last);
    }

    public static void ValidateSheetName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new WriteValidationException("sheet name is empty");
        if (name.Length > 31) throw new WriteValidationException($"sheet name '{name}' is longer than 31 characters");
        var bad = name.IndexOfAny(ForbiddenSheetNameChars);
        if (bad >= 0) throw new WriteValidationException($"sheet name '{name}' contains '{name[bad]}'; none of : \\ / ? * [ ] is allowed");
        if (name[0] == '\'' || name[^1] == '\'') throw new WriteValidationException($"sheet name '{name}' may not start or end with an apostrophe");
    }

    /// <summary>Only <c>#RRGGBB</c>.</summary>
    public static Color ParseHexColor(string? hex)
    {
        if (hex is null || !HexColorRegex().IsMatch(hex))
            throw new WriteValidationException($"colour '{hex}' is not #RRGGBB");
        return Color.FromArgb(
            int.Parse(hex.AsSpan(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            int.Parse(hex.AsSpan(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            int.Parse(hex.AsSpan(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
    }
}
