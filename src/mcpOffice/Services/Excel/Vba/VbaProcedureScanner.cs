using System.Text.RegularExpressions;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static partial class VbaProcedureScanner
{
    [GeneratedRegex(
        @"^\s*(?<scope>Public|Private|Friend)?\s*(Static\s+)?(?<kind>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+(?<name>\w+)\s*\((?<params>[^)]*)\)(\s+As\s+(?<ret>\w+))?",
        RegexOptions.IgnoreCase)]
    private static partial Regex ProcOpenRegex();

    [GeneratedRegex(@"^\s*End\s+(Sub|Function|Property)\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex ProcCloseRegex();

    public static IReadOnlyList<ScannedProcedure> Scan(string moduleKind, string moduleName, IReadOnlyList<CleanedLine> lines)
    {
        var procs = new List<ScannedProcedure>();

        for (int i = 0; i < lines.Count; i++)
        {
            var open = ProcOpenRegex().Match(lines[i].Text);
            if (!open.Success) continue;

            int startLine = lines[i].LineNumber;
            int bodyStartIdx = i + 1;
            int closeIdx = -1;
            for (int j = i + 1; j < lines.Count; j++)
            {
                if (ProcCloseRegex().IsMatch(lines[j].Text)) { closeIdx = j; break; }
            }
            int endLine = closeIdx >= 0 ? lines[closeIdx].LineNumber : lines[^1].LineNumber;

            var name = open.Groups["name"].Value;
            var kindRaw = open.Groups["kind"].Value;
            var kind = NormalizeKind(kindRaw);
            var scope = open.Groups["scope"].Success ? open.Groups["scope"].Value : null;
            var returnType = open.Groups["ret"].Success ? open.Groups["ret"].Value : null;
            var paramList = ParseParameters(open.Groups["params"].Value);

            var (isEvent, target) = ClassifyEventHandler(moduleKind, name);

            var proc = new ExcelVbaProcedure(
                Name: name,
                FullyQualifiedName: $"{moduleName}.{name}",
                Kind: kind,
                Scope: scope,
                Parameters: paramList,
                ReturnType: returnType,
                LineStart: startLine,
                LineEnd: endLine,
                IsEventHandler: isEvent,
                EventTarget: target);

            procs.Add(new ScannedProcedure(proc, bodyStartIdx, closeIdx >= 0 ? closeIdx - 1 : lines.Count - 1));
            i = closeIdx >= 0 ? closeIdx : lines.Count;
        }

        return procs;
    }

    private static string NormalizeKind(string raw)
    {
        var collapsed = Regex.Replace(raw, @"\s+", "");
        return collapsed switch
        {
            "Sub" => "Sub",
            "Function" => "Function",
            "PropertyGet" => "PropertyGet",
            "PropertyLet" => "PropertyLet",
            "PropertySet" => "PropertySet",
            _ => collapsed
        };
    }

    private static IReadOnlyList<ExcelVbaParameter> ParseParameters(string paramText)
    {
        if (string.IsNullOrWhiteSpace(paramText)) return [];
        var parts = SplitOnTopLevelCommas(paramText);
        var list = new List<ExcelVbaParameter>(parts.Count);
        foreach (var part in parts)
        {
            var p = part.Trim();
            if (p.Length == 0) continue;

            bool optional = false;
            bool byRef = true; // VBA default is ByRef
            var modPattern = @"^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?";
            var m = Regex.Match(p, modPattern, RegexOptions.IgnoreCase);
            if (m.Success)
            {
                if (m.Value.IndexOf("Optional", StringComparison.OrdinalIgnoreCase) >= 0) optional = true;
                if (m.Value.IndexOf("ByVal", StringComparison.OrdinalIgnoreCase) >= 0) byRef = false;
                p = p[m.Length..];
            }

            string name;
            string? type = null;
            string? defaultValue = null;

            var eq = FindTopLevelEquals(p);
            if (eq >= 0)
            {
                defaultValue = p[(eq + 1)..].Trim();
                p = p[..eq].Trim();
            }

            var asIdx = Regex.Match(p, @"\s+As\s+", RegexOptions.IgnoreCase);
            if (asIdx.Success)
            {
                name = p[..asIdx.Index].Trim();
                type = p[(asIdx.Index + asIdx.Length)..].Trim();
            }
            else
            {
                name = p.Trim();
            }

            list.Add(new ExcelVbaParameter(name, type, byRef, optional, defaultValue));
        }
        return list;
    }

    private static List<string> SplitOnTopLevelCommas(string s)
    {
        var parts = new List<string>();
        int depth = 0;
        int start = 0;
        for (int i = 0; i < s.Length; i++)
        {
            var c = s[i];
            if (c == '(' || c == '[') depth++;
            else if (c == ')' || c == ']') depth--;
            else if (c == ',' && depth == 0)
            {
                parts.Add(s[start..i]);
                start = i + 1;
            }
        }
        parts.Add(s[start..]);
        return parts;
    }

    private static int FindTopLevelEquals(string s)
    {
        int depth = 0;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '(' || s[i] == '[') depth++;
            else if (s[i] == ')' || s[i] == ']') depth--;
            else if (s[i] == '=' && depth == 0) return i;
        }
        return -1;
    }

    private static (bool IsEvent, string? Target) ClassifyEventHandler(string moduleKind, string name)
    {
        if (moduleKind == "standardModule") return (false, null);
        var idx = name.IndexOf('_');
        if (idx <= 0 || idx == name.Length - 1) return (false, null);
        return (true, name[..idx]);
    }
}
