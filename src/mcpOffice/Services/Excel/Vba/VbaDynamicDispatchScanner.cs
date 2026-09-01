using System.Text;
using System.Text.RegularExpressions;

namespace McpOffice.Services.Excel.Vba;

/// <summary>
/// Finds procedures invoked by name at runtime (`Application.OnTime/OnKey/Run`, `.OnAction =`,
/// `CallByName`). The target is reported verbatim when it is a string literal, null otherwise.
/// </summary>
internal static partial class VbaDynamicDispatchScanner
{
    public sealed record DynamicDispatch(string Module, string Procedure, int Line, string Api, string? TargetLiteral);

    // (regex, api, positional index of the target argument, its named-argument name). null index = the
    // whole remainder is the target expression (property assignment).
    private static readonly (Regex Pattern, string Api, int? ArgIndex, string? NamedArg)[] Apis =
    [
        (OnTimeRegex(), "OnTime", 1, "Procedure"),
        (OnKeyRegex(), "OnKey", 1, "Procedure"),
        (RunRegex(), "Run", 0, "Macro"),
        (OnActionRegex(), "OnAction", null, null),
        (CallByNameRegex(), "CallByName", 1, "ProcName"),
    ];

    [GeneratedRegex(@"\bApplication\.OnTime\b", RegexOptions.IgnoreCase)]
    private static partial Regex OnTimeRegex();

    [GeneratedRegex(@"\bApplication\.OnKey\b", RegexOptions.IgnoreCase)]
    private static partial Regex OnKeyRegex();

    // `Application.Run …` anywhere, or a statement-level `Run "x"` / `Call Run(...)` — but not `Run = 5`
    // and not `wsh.Run` (WScript.Shell).
    [GeneratedRegex(@"\bApplication\.Run\b|^\s*(?:Call\s+)?Run(?=\s*\(|\s+[^=\s])", RegexOptions.IgnoreCase)]
    private static partial Regex RunRegex();

    // ponytail: also matches `If x.OnAction = "…"` comparisons — over-approximates reachability, the safe side.
    [GeneratedRegex(@"\.OnAction\s*=", RegexOptions.IgnoreCase)]
    private static partial Regex OnActionRegex();

    [GeneratedRegex(@"\bCallByName\b", RegexOptions.IgnoreCase)]
    private static partial Regex CallByNameRegex();

    [GeneratedRegex(@"^(?<name>\w+)\s*:=\s*(?<value>.*)$", RegexOptions.Singleline)]
    private static partial Regex NamedArgRegex();

    [GeneratedRegex(@"^""(?<lit>(?:[^""]|"""")*)""$")]
    private static partial Regex StringLiteralRegex();

    public static IReadOnlyList<DynamicDispatch> Scan(string moduleName, IReadOnlyList<CleanedLine> lines, IReadOnlyList<ScannedProcedure> procs)
    {
        var result = new List<DynamicDispatch>();
        foreach (var sp in procs)
        {
            for (int i = sp.CleanedLineStartIndex; i <= sp.CleanedLineEndIndex && i < lines.Count; i++)
            {
                var line = lines[i];
                string? code = null;   // OriginalText minus trailing comment; literals intact
                foreach (var (pattern, api, argIndex, namedArg) in Apis)
                {
                    // Text has literals blanked and comments removed: a hit there is a real call, not prose.
                    if (!pattern.IsMatch(line.Text)) continue;
                    code ??= StripComment(line.OriginalText);
                    var m = pattern.Match(code);
                    if (!m.Success) continue;

                    var rest = code[(m.Index + m.Length)..];
                    var target = argIndex is null ? rest.Trim() : TargetArgument(rest, argIndex.Value, namedArg!);
                    if (target is null) continue;   // no target argument at all (e.g. OnKey reset form)

                    var lit = StringLiteralRegex().Match(target);
                    result.Add(new DynamicDispatch(moduleName, sp.Procedure.Name, line.LineNumber, api,
                        lit.Success ? lit.Groups["lit"].Value.Replace("\"\"", "\"") : null));
                }
            }
        }
        return result;
    }

    private static string? TargetArgument(string argText, int index, string namedArg)
    {
        var args = SplitArguments(argText);
        foreach (var arg in args)
        {
            var named = NamedArgRegex().Match(arg);
            if (named.Success && named.Groups["name"].Value.Equals(namedArg, StringComparison.OrdinalIgnoreCase))
                return named.Groups["value"].Value.Trim();
        }
        if (index >= args.Count) return null;
        var positional = args[index];
        return NamedArgRegex().IsMatch(positional) ? null : positional;   // a different named arg sits there
    }

    /// <summary>Top-level comma split; strips one pair of wrapping parens (`Run("x", 1)` → `"x", 1`).</summary>
    private static List<string> SplitArguments(string text)
    {
        var s = text.Trim();
        if (s.StartsWith('(') && MatchingClose(s, 0) == s.Length - 1) s = s[1..^1];

        var args = new List<string>();
        var current = new StringBuilder();
        int depth = 0;
        bool inString = false;
        foreach (var c in s)
        {
            if (c == '"') inString = !inString;
            else if (!inString)
            {
                if (c == '(') depth++;
                else if (c == ')') depth--;
                else if (c == ',' && depth == 0) { args.Add(current.ToString().Trim()); current.Clear(); continue; }
            }
            current.Append(c);
        }
        args.Add(current.ToString().Trim());
        return args;
    }

    private static int MatchingClose(string s, int openIndex)
    {
        int depth = 0;
        bool inString = false;
        for (int i = openIndex; i < s.Length; i++)
        {
            var c = s[i];
            if (c == '"') inString = !inString;
            else if (inString) continue;
            else if (c == '(') depth++;
            else if (c == ')' && --depth == 0) return i;
        }
        return -1;
    }

    // OriginalText keeps the trailing comment the cleaner removed; drop it without touching literals.
    private static string StripComment(string raw)
    {
        bool inString = false;
        for (int i = 0; i < raw.Length; i++)
        {
            if (raw[i] == '"') inString = !inString;
            else if (raw[i] == '\'' && !inString) return raw[..i];
        }
        return raw;
    }
}
