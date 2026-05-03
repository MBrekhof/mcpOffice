using System.Text;

namespace McpOffice.Services.Excel.Vba;

internal static class VbaLineCleaner
{
    private const string StringSentinel = "<STR>";

    public static IReadOnlyList<CleanedLine> Clean(string source)
    {
        if (string.IsNullOrEmpty(source)) return [];
        var rawLines = source.Replace("\r\n", "\n").Split('\n');
        var result = new List<CleanedLine>(rawLines.Length);

        var pending = new StringBuilder();
        var pendingOriginal = new StringBuilder();
        int? pendingStart = null;

        for (int i = 0; i < rawLines.Length; i++)
        {
            var raw = rawLines[i];
            var cleaned = CleanSingleLine(raw);

            if (pending.Length == 0)
            {
                pendingStart = i + 1;
            }

            var endsWithContinuation = EndsWithContinuation(raw);

            if (endsWithContinuation)
            {
                // The continuation `_` may have been inside a comment that CleanSingleLine
                // already stripped. Only strip from `cleaned` if it actually still ends with `_`.
                pending.Append(EndsWithContinuation(cleaned) ? StripTrailingContinuation(cleaned) : cleaned);
                pending.Append(' ');
                pendingOriginal.Append(StripTrailingContinuation(raw));
                pendingOriginal.Append(' ');
                continue;
            }

            pending.Append(cleaned);
            pendingOriginal.Append(raw);

            result.Add(new CleanedLine(pendingStart ?? (i + 1), pending.ToString(), pendingOriginal.ToString()));
            pending.Clear();
            pendingOriginal.Clear();
            pendingStart = null;
        }

        if (pending.Length > 0)
        {
            result.Add(new CleanedLine(pendingStart ?? rawLines.Length, pending.ToString(), pendingOriginal.ToString()));
        }

        return result;
    }

    private static string CleanSingleLine(string raw)
    {
        var trimmed = raw.TrimStart();
        if (trimmed.Equals("Rem", StringComparison.OrdinalIgnoreCase) ||
            (trimmed.Length >= 4 &&
             trimmed.StartsWith("Rem", StringComparison.OrdinalIgnoreCase) &&
             char.IsWhiteSpace(trimmed[3])))
        {
            return new string(' ', raw.Length - trimmed.Length);
        }

        var sb = new StringBuilder(raw.Length);
        for (int i = 0; i < raw.Length; i++)
        {
            char c = raw[i];

            if (c == '"')
            {
                sb.Append('"').Append(StringSentinel).Append('"');
                int j = i + 1;
                while (j < raw.Length)
                {
                    if (raw[j] == '"')
                    {
                        if (j + 1 < raw.Length && raw[j + 1] == '"') { j += 2; continue; }
                        break;
                    }
                    j++;
                }
                i = j;
                continue;
            }

            if (c == '\'') return sb.ToString();

            sb.Append(c);
        }
        return sb.ToString();
    }

    private static bool EndsWithContinuation(string raw)
    {
        var trimmed = raw.TrimEnd();
        if (trimmed.Length < 2) return false;
        if (trimmed[^1] != '_') return false;
        return char.IsWhiteSpace(trimmed[^2]);
    }

    private static string StripTrailingContinuation(string s)
    {
        var trimmed = s.TrimEnd();
        return trimmed[..^1];
    }
}
