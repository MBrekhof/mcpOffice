using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace McpOffice.Services.Excel.Vba;

/// <summary>
/// Procedure identity for cross-workbook comparison: the normalised body (not the name, so a
/// renamed copy still groups). Pure.
/// </summary>
internal static partial class VbaProcedureHasher
{
    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();

    /// <summary>
    /// Body lines of a procedure with trailing comments removed (string literals kept — two bodies
    /// that differ only in a literal are near-duplicates, not identical), whitespace collapsed,
    /// case-folded (VBA is case-insensitive), <c>Attribute</c> and blank lines dropped.
    /// </summary>
    public static IReadOnlyList<string> Normalize(IReadOnlyList<CleanedLine> lines, int startIndex, int endIndex)
    {
        var result = new List<string>();
        for (int i = Math.Max(0, startIndex); i <= endIndex && i < lines.Count; i++)
        {
            var text = WhitespaceRegex().Replace(StripTrailingComment(lines[i].OriginalText).Trim(), " ");
            if (text.Length == 0) continue;
            if (text.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase)) continue;
            result.Add(text.ToLowerInvariant());
        }
        return result;
    }

    /// <summary>Drops a trailing `'` comment that is not inside a string literal.</summary>
    public static string StripTrailingComment(string line)
    {
        bool inStr = false;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"') inStr = !inStr;
            else if (line[i] == '\'' && !inStr) return line[..i];
        }
        return line;
    }

    public static string Hash(IReadOnlyList<string> normalized)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(string.Join('\n', normalized)));
        return Convert.ToHexStringLower(bytes);
    }

    /// <summary>
    /// 2·|common lines (multiset)| / (|a| + |b|). Two empty bodies are identical (1.0).
    /// ponytail: line-multiset similarity, not LCS — upgrade if reordered bodies turn out to matter.
    /// </summary>
    public static double Similarity(IReadOnlyList<string> a, IReadOnlyList<string> b)
    {
        if (a.Count == 0 && b.Count == 0) return 1.0;
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var line in a) counts[line] = counts.GetValueOrDefault(line) + 1;
        int common = 0;
        foreach (var line in b)
        {
            if (counts.TryGetValue(line, out var n) && n > 0)
            {
                counts[line] = n - 1;
                common++;
            }
        }
        return 2.0 * common / (a.Count + b.Count);
    }
}
