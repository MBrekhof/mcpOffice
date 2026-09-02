using System.Text.RegularExpressions;

namespace McpOffice.Services.Excel.Vba;

/// <summary>
/// Resolves each Range/Cells/Columns/Rows/UsedRange/[A1] site in a procedure to a sheet and a
/// target, with read/write mode. Regex on comment-stripped source lines, With-block and
/// one-assignment alias tracking per procedure. Never guesses a sheet: ActiveSheet and
/// unqualified access outside a sheet module come back unresolved with a reason.
/// </summary>
internal static partial class VbaSheetAccessResolver
{
    public sealed record SheetName(string Name, string? CodeName);
    public sealed record DefinedName(string Name, string? Scope, string? RefersTo);
    public sealed record AccessSite(
        string Module, string Procedure, int Line,
        string? SheetName, string? CodeName,
        string TargetKind, string? Address, string? DefinedNameRef,
        string Mode, string? UnresolvedReason);

    private static readonly StringComparer Ci = StringComparer.OrdinalIgnoreCase;
    private static readonly string[] SiteKeywords = ["Range", "Cells", "Columns", "Rows", "UsedRange"];

    [GeneratedRegex(@"\b(?<kw>Range|Cells|Columns|Rows|UsedRange)\b|\[(?<bracket>[A-Za-z_$][^\]\s]*)\]", RegexOptions.IgnoreCase)]
    private static partial Regex SiteRegex();

    [GeneratedRegex(@"^\s*With\s+(?<expr>.+?)\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex WithRegex();

    [GeneratedRegex(@"^\s*End\s+With\b", RegexOptions.IgnoreCase)]
    private static partial Regex EndWithRegex();

    [GeneratedRegex(@"^\s*Set\s+(?<name>[A-Za-z_]\w*)\s*=\s*(?<expr>.+?)\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex SetRegex();

    [GeneratedRegex(@"^(?<api>Worksheets|Sheets)\s*\((?<arg>.*)\)$", RegexOptions.IgnoreCase)]
    private static partial Regex SheetsCallRegex();

    [GeneratedRegex(@"^\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?$")]
    private static partial Regex A1Regex();

    [GeneratedRegex(@"^\$?[A-Za-z]{1,3}(?::\$?[A-Za-z]{1,3})?$")]
    private static partial Regex ColumnRefRegex();

    [GeneratedRegex(@"^\$?\d{1,7}(?::\$?\d{1,7})?$")]
    private static partial Regex RowRefRegex();

    [GeneratedRegex(@"^=?(?:'(?<qsheet>[^']+)'|(?<sheet>[^'!]+))!(?<addr>\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)$")]
    private static partial Regex RefersToRegex();

    [GeneratedRegex(@"^\s*(If|ElseIf|While|Until|Do|Select|Case|Loop)\b", RegexOptions.IgnoreCase)]
    private static partial Regex ConditionLineRegex();

    // Member chain right after a site that mutates the target. `.Copy <dest>` writes the destination,
    // handled separately; `.Copy` alone (to clipboard) is a read.
    [GeneratedRegex(@"^(?:\.\w+(?:\([^()]*\))?)*?\.(?:Clear|ClearContents|ClearFormats|ClearComments|Delete|Insert|AutoFilter|Sort|Merge|UnMerge|PasteSpecial|Paste|FillDown|FillRight|AutoFill|RemoveDuplicates|Replace)\b", RegexOptions.IgnoreCase)]
    private static partial Regex MutatingMemberRegex();

    // Leading-dot members inside a `With <range>` block that write the With target.
    [GeneratedRegex(@"^\s*\.(?:Value2?|Formula(?:R1C1)?|Text|NumberFormat|Interior|Font)\b[^=<>]*=(?!=)|^\s*\.(?:Clear|ClearContents|ClearFormats|Delete|Insert|Sort|AutoFilter|Merge|PasteSpecial)\b", RegexOptions.IgnoreCase)]
    private static partial Regex WithMemberWriteRegex();

    [GeneratedRegex(@"\.Copy\s+(?:Destination:=)?\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex CopyDestinationPrefixRegex();

    private sealed record SheetRes(string? Name, string? CodeName, string? Reason)
    {
        public static readonly SheetRes ActiveSheet = new(null, null, "activeSheet");
        public static readonly SheetRes Dynamic = new(null, null, "dynamicSheet");
        public static readonly SheetRes Unknown = new(null, null, "unknownSheet");
        public static readonly SheetRes Reassigned = new(null, null, "aliasReassigned");
    }

    private sealed record WithFrame(SheetRes Sheet, (string Kind, string? Address, string? Name)? Target);

    private sealed class Context(
        string module, string moduleKind, string procedure,
        IReadOnlyList<SheetName> sheets, IReadOnlyList<DefinedName> definedNames)
    {
        public string Module { get; } = module;
        public string Procedure { get; } = procedure;
        public SheetRes? OwnSheet { get; } =
            moduleKind == "documentModule"
                ? sheets.Where(s => s.CodeName is not null && Ci.Equals(s.CodeName, module))
                        .Select(s => new SheetRes(s.Name, s.CodeName, null)).FirstOrDefault()
                : null;
        public IReadOnlyList<SheetName> Sheets { get; } = sheets;
        public IReadOnlyList<DefinedName> DefinedNames { get; } = definedNames;
        public Stack<WithFrame> With { get; } = new();
        public Dictionary<string, SheetRes> Aliases { get; } = new(Ci);
    }

    public static IReadOnlyList<AccessSite> Resolve(
        string moduleName, string moduleKind, IReadOnlyList<CleanedLine> lines, IReadOnlyList<ScannedProcedure> procs,
        IReadOnlyList<SheetName> sheets, IReadOnlyList<DefinedName> definedNames)
    {
        var sites = new List<AccessSite>();
        foreach (var sp in procs)
        {
            var ctx = new Context(moduleName, moduleKind, sp.Procedure.Name, sheets, definedNames);
            for (int i = sp.CleanedLineStartIndex; i <= sp.CleanedLineEndIndex && i < lines.Count; i++)
            {
                var line = lines[i];
                var text = VbaProcedureHasher.StripTrailingComment(line.OriginalText);
                if (text.Trim().Length == 0) continue;
                ProcessLine(ctx, text, line.LineNumber, sites);
            }
        }
        return sites;
    }

    private static void ProcessLine(Context ctx, string text, int lineNumber, List<AccessSite> sites)
    {
        if (EndWithRegex().IsMatch(text))
        {
            if (ctx.With.Count > 0) ctx.With.Pop();
            return;
        }

        var with = WithRegex().Match(text);
        if (with.Success)
        {
            var expr = with.Groups["expr"].Value;
            // `With Worksheets("X").Range("A1")` — the target is a range: record it (read) and
            // remember it so leading-dot writes inside the block hit the same target.
            var inner = FindSites(ctx, expr, lineNumber, isWithExpression: true);
            if (inner.Count > 0)
            {
                var t = inner[^1];
                sites.Add(t);
                ctx.With.Push(new WithFrame(new SheetRes(t.SheetName, t.CodeName, t.UnresolvedReason), (t.TargetKind, t.Address, t.DefinedNameRef)));
            }
            else
            {
                ctx.With.Push(new WithFrame(ResolveQualifier(ctx, expr.Trim()), null));
            }
            return;
        }

        var set = SetRegex().Match(text);
        if (set.Success)
        {
            var name = set.Groups["name"].Value;
            var res = ResolveQualifier(ctx, set.Groups["expr"].Value.Trim());
            if (res.Name is not null || res.Reason == "dynamicSheet" || res.Reason == "unknownSheet")
            {
                if (ctx.Aliases.TryGetValue(name, out var prev) && !SameSheet(prev, res))
                    ctx.Aliases[name] = SheetRes.Reassigned;
                else if (!ctx.Aliases.ContainsKey(name) || ctx.Aliases[name].Reason != "aliasReassigned")
                    ctx.Aliases[name] = res;
            }
            // `Set rng = Worksheets("X").Range("A1")` also reads that range.
            sites.AddRange(FindSites(ctx, set.Groups["expr"].Value, lineNumber, isWithExpression: false));
            return;
        }

        // Leading-dot write inside a `With <range>` block.
        if (ctx.With.Count > 0 && ctx.With.Peek().Target is { } target && WithMemberWriteRegex().IsMatch(text))
        {
            var frame = ctx.With.Peek();
            sites.Add(new AccessSite(ctx.Module, ctx.Procedure, lineNumber, frame.Sheet.Name, frame.Sheet.CodeName,
                target.Kind, target.Address, target.Name, "write", frame.Sheet.Name is null ? frame.Sheet.Reason : null));
            return;
        }

        sites.AddRange(FindSites(ctx, text, lineNumber, isWithExpression: false));
    }

    private static bool SameSheet(SheetRes a, SheetRes b) =>
        a.Name is not null && b.Name is not null ? Ci.Equals(a.Name, b.Name) : a.Reason == b.Reason && a.Name == b.Name;

    private static List<AccessSite> FindSites(Context ctx, string text, int lineNumber, bool isWithExpression)
    {
        var result = new List<AccessSite>();
        // Structure is detected on a copy with string-literal contents masked (same length, so
        // indices line up); literal values are sliced from the original at the same positions.
        var masked = MaskStrings(text);
        int assignIndex = isWithExpression ? -1 : FindAssignmentIndex(masked);
        int skipUntil = -1;

        foreach (Match m in SiteRegex().Matches(masked))
        {
            if (m.Index < skipUntil) continue;   // nested inside another site's argument list

            var sheet = QualifierBefore(ctx, masked, text, m.Index, out var skip, out var siteStart);
            if (skip) continue;

            string kind; string? address; string? definedName; SheetRes? nameSheet;
            int end = m.Index + m.Length;

            if (m.Groups["bracket"].Success)
            {
                var g = m.Groups["bracket"];
                (kind, address, definedName, nameSheet) = ClassifyLiteral(ctx, text.Substring(g.Index, g.Length));
            }
            else
            {
                string? args = null;
                if (end < masked.Length && masked[end] == '(')
                {
                    var close = FindMatchingParen(masked, end);
                    if (close < 0) continue;
                    args = text[(end + 1)..close];
                    skipUntil = close;
                    end = close + 1;
                }
                (kind, address, definedName, nameSheet) = ClassifyTarget(ctx, m.Groups["kw"].Value, args);
                if (kind == "") continue;
            }
            if (nameSheet is not null && (sheet.Name is null || nameSheet.Reason == "unknownName")) sheet = nameSheet;

            var mode = "read";
            if (!isWithExpression)
            {
                var after = masked[end..];
                var before = masked[..siteStart];
                if (assignIndex > m.Index || MutatingMemberRegex().IsMatch(after) || CopyDestinationPrefixRegex().IsMatch(before))
                    mode = "write";
            }

            result.Add(new AccessSite(ctx.Module, ctx.Procedure, lineNumber, sheet.Name, sheet.CodeName,
                kind, address, definedName, mode, sheet.Name is null ? sheet.Reason : null));
        }
        return result;
    }

    /// <summary>
    /// Sheet for the site at <paramref name="index"/> from what precedes it. <paramref name="skip"/>
    /// when nested inside another range expression; <paramref name="siteStart"/> is the start of
    /// the whole qualifier chain (`ThisWorkbook.Worksheets("X").`), used for the `.Copy dest` check.
    /// </summary>
    private static SheetRes QualifierBefore(Context ctx, string masked, string text, int index, out bool skip, out int siteStart)
    {
        skip = false;
        siteStart = index;
        int i = index - 1;
        if (i < 0 || masked[i] != '.')
            return ctx.OwnSheet ?? SheetRes.ActiveSheet;

        var segment = SegmentBefore(masked, i - 1, out var segStart);
        if (segment.Length == 0)
        {
            // Leading dot: With member.
            siteStart = i;
            if (ctx.With.Count == 0) return SheetRes.Dynamic;
            var frame = ctx.With.Peek();
            if (frame.Target is not null) { skip = true; return frame.Sheet; }   // `.Cells(1,1)` inside With <range>
            return frame.Sheet;
        }

        siteStart = ChainStart(masked, segStart);
        var head = segment.Contains('(') ? segment[..segment.IndexOf('(')] : segment;
        if (SiteKeywords.Contains(head, Ci)) { skip = true; return SheetRes.Dynamic; }   // Range(...).Cells(...): outer already counted
        if (Ci.Equals(head, "Item") || Ci.Equals(head, "Offset") || Ci.Equals(head, "Resize") || Ci.Equals(head, "End")) { skip = true; return SheetRes.Dynamic; }

        return ResolveQualifier(ctx, text[segStart..i]);
    }

    /// <summary>Walks back over `Ident.` / `Ident(...).` segments to the start of the chain.</summary>
    private static int ChainStart(string masked, int segStart)
    {
        int s = segStart;
        while (s > 0 && masked[s - 1] == '.')
        {
            var seg = SegmentBefore(masked, s - 2, out var st);
            if (seg.Length == 0) break;
            s = st;
        }
        return s;
    }

    /// <summary>Same length as <paramref name="text"/>, with the contents of "…" literals replaced by 'x'.</summary>
    private static string MaskStrings(string text)
    {
        var chars = text.ToCharArray();
        bool inStr = false;
        for (int i = 0; i < chars.Length; i++)
        {
            if (chars[i] == '"') inStr = !inStr;
            else if (inStr) chars[i] = 'x';
        }
        return new string(chars);
    }

    /// <summary>Identifier or `Ident(...)` immediately before position <paramref name="i"/> (the char before a dot).</summary>
    private static string SegmentBefore(string text, int i, out int start)
    {
        start = i + 1;
        if (i < 0) return "";
        int j = i;
        if (text[j] == ')')
        {
            int open = FindMatchingOpen(text, j);
            if (open < 0) return "";
            j = open - 1;
        }
        int identEnd = j;
        while (j >= 0 && (char.IsLetterOrDigit(text[j]) || text[j] == '_')) j--;
        if (identEnd == j) return "";
        start = j + 1;
        return text[start..(i + 1)];
    }

    private static SheetRes ResolveQualifier(Context ctx, string segment)
    {
        // Strip harmless workbook prefixes: ThisWorkbook.Sheets("X"), ActiveWorkbook.Worksheets(1), Workbooks("B").Sheets("X").
        var lastDot = LastTopLevelDot(segment);
        var tail = lastDot >= 0 ? segment[(lastDot + 1)..] : segment;

        var call = SheetsCallRegex().Match(tail.Trim());
        if (call.Success)
        {
            var arg = call.Groups["arg"].Value.Trim();
            if (arg.Length >= 2 && arg[0] == '"' && arg[^1] == '"')
            {
                var name = arg[1..^1];
                var s = ctx.Sheets.FirstOrDefault(x => Ci.Equals(x.Name, name));
                return s is null ? SheetRes.Unknown : new SheetRes(s.Name, s.CodeName, null);
            }
            if (int.TryParse(arg, out var n))
            {
                return n >= 1 && n <= ctx.Sheets.Count
                    ? new SheetRes(ctx.Sheets[n - 1].Name, ctx.Sheets[n - 1].CodeName, null)
                    : SheetRes.Unknown;
            }
            return SheetRes.Dynamic;
        }

        var ident = tail.Trim();
        if (Ci.Equals(ident, "ActiveSheet")) return SheetRes.ActiveSheet;
        if (Ci.Equals(ident, "Me")) return ctx.OwnSheet ?? SheetRes.ActiveSheet;
        if (Ci.Equals(ident, "Application") || Ci.Equals(ident, "ThisWorkbook") || Ci.Equals(ident, "ActiveWorkbook")) return SheetRes.ActiveSheet;
        var byCode = ctx.Sheets.FirstOrDefault(s => s.CodeName is not null && Ci.Equals(s.CodeName, ident));
        if (byCode is not null) return new SheetRes(byCode.Name, byCode.CodeName, null);
        if (ctx.Aliases.TryGetValue(ident, out var alias)) return alias;
        return SheetRes.Dynamic;
    }

    private static int LastTopLevelDot(string s)
    {
        int depth = 0, last = -1; bool inStr = false;
        for (int i = 0; i < s.Length; i++)
        {
            var c = s[i];
            if (c == '"') inStr = !inStr;
            if (inStr) continue;
            if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == '.' && depth == 0) last = i;
        }
        return last;
    }

    private static (string Kind, string? Address, string? DefinedName, SheetRes? Sheet) ClassifyTarget(Context ctx, string kw, string? args)
    {
        var trimmed = args?.Trim();
        switch (kw.ToLowerInvariant())
        {
            case "usedrange":
                return ("wholeSheet", null, null, null);
            case "range":
                if (trimmed is null || trimmed.Length == 0) return ("wholeSheet", null, null, null);
                if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[^1] == '"' && !trimmed[1..^1].Contains('"'))
                    return ClassifyLiteral(ctx, trimmed[1..^1]);
                return ("dynamicCells", null, null, null);   // Range(var), Range(Cells(..), Cells(..))
            case "cells":
                if (trimmed is null || trimmed.Length == 0) return ("wholeSheet", null, null, null);
                {
                    var parts = SplitTopLevel(trimmed);
                    if (parts.Count == 2 && int.TryParse(parts[0].Trim(), out var r) && int.TryParse(parts[1].Trim(), out var c) && r > 0 && c > 0)
                        return ("range", ColumnLetters(c) + r, null, null);
                    return ("dynamicCells", null, null, null);
                }
            case "columns":
                if (trimmed is null || trimmed.Length == 0) return ("wholeSheet", null, null, null);
                if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[^1] == '"' && ColumnRefRegex().IsMatch(trimmed[1..^1]))
                    return ("column", trimmed[1..^1].Replace("$", ""), null, null);
                if (int.TryParse(trimmed, out var col) && col > 0) return ("column", ColumnLetters(col), null, null);
                return ("dynamicCells", null, null, null);
            case "rows":
                if (trimmed is null || trimmed.Length == 0) return ("wholeSheet", null, null, null);
                if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[^1] == '"' && RowRefRegex().IsMatch(trimmed[1..^1]))
                    return ("row", trimmed[1..^1].Replace("$", ""), null, null);
                if (int.TryParse(trimmed, out var row) && row > 0) return ("row", row.ToString(), null, null);
                return ("dynamicCells", null, null, null);
        }
        return ("", null, null, null);
    }

    /// <summary>A1 address → range; a defined name → definedName (sheet/address from refersTo); else unknownName.</summary>
    private static (string Kind, string? Address, string? DefinedName, SheetRes? Sheet) ClassifyLiteral(Context ctx, string literal)
    {
        var s = literal.Trim();
        if (A1Regex().IsMatch(s)) return ("range", s.Replace("$", "").ToUpperInvariant(), null, null);
        if (s.Contains('!'))
        {
            // Sheet-qualified literal: Range("Data!A1")
            var rt = RefersToRegex().Match(s);
            if (rt.Success)
            {
                var sheetName = rt.Groups["qsheet"].Success ? rt.Groups["qsheet"].Value : rt.Groups["sheet"].Value;
                var sheet = ctx.Sheets.FirstOrDefault(x => Ci.Equals(x.Name, sheetName));
                return ("range", rt.Groups["addr"].Value.Replace("$", "").ToUpperInvariant(), null,
                    sheet is null ? SheetRes.Unknown : new SheetRes(sheet.Name, sheet.CodeName, null));
            }
        }
        var dn = ctx.DefinedNames.FirstOrDefault(d => Ci.Equals(d.Name, s));
        if (dn is not null)
        {
            var rt = dn.RefersTo is null ? Match.Empty : RefersToRegex().Match(dn.RefersTo.Trim());
            if (rt.Success)
            {
                var sheetName = rt.Groups["qsheet"].Success ? rt.Groups["qsheet"].Value : rt.Groups["sheet"].Value;
                var sheet = ctx.Sheets.FirstOrDefault(x => Ci.Equals(x.Name, sheetName));
                return ("definedName", rt.Groups["addr"].Value.Replace("$", "").ToUpperInvariant(), dn.Name,
                    sheet is null ? SheetRes.Unknown : new SheetRes(sheet.Name, sheet.CodeName, null));
            }
            return ("definedName", null, dn.Name, SheetRes.Unknown);   // named formula or constant
        }
        return ("range", null, null, new SheetRes(null, null, "unknownName"));
    }

    /// <summary>Index of the assignment `=` at paren depth 0, or -1 when the line is a condition / has none.</summary>
    private static int FindAssignmentIndex(string text)
    {
        if (ConditionLineRegex().IsMatch(text)) return -1;
        int depth = 0; bool inStr = false;
        for (int i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (c == '"') { inStr = !inStr; continue; }
            if (inStr) continue;
            if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == '=' && depth == 0)
            {
                if (i > 0 && (text[i - 1] == ':' || text[i - 1] == '<' || text[i - 1] == '>')) continue;   // :=  <=  >=
                if (i + 1 < text.Length && text[i + 1] == '=') continue;
                return i;
            }
        }
        return -1;
    }

    private static List<string> SplitTopLevel(string s)
    {
        var parts = new List<string>(); int depth = 0, start = 0; bool inStr = false;
        for (int i = 0; i < s.Length; i++)
        {
            var c = s[i];
            if (c == '"') inStr = !inStr;
            if (inStr) continue;
            if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == ',' && depth == 0) { parts.Add(s[start..i]); start = i + 1; }
        }
        parts.Add(s[start..]);
        return parts;
    }

    private static int FindMatchingParen(string text, int openIndex)
    {
        int depth = 0; bool inStr = false;
        for (int i = openIndex; i < text.Length; i++)
        {
            var c = text[i];
            if (c == '"') inStr = !inStr;
            if (inStr) continue;
            if (c == '(') depth++;
            else if (c == ')' && --depth == 0) return i;
        }
        return -1;
    }

    private static int FindMatchingOpen(string text, int closeIndex)
    {
        int depth = 0; bool inStr = false;
        for (int i = closeIndex; i >= 0; i--)
        {
            var c = text[i];
            if (c == '"') inStr = !inStr;
            if (inStr) continue;
            if (c == ')') depth++;
            else if (c == '(' && --depth == 0) return i;
        }
        return -1;
    }

    private static string ColumnLetters(int c)
    {
        var s = "";
        while (c > 0) { c--; s = (char)('A' + c % 26) + s; c /= 26; }
        return s;
    }
}
