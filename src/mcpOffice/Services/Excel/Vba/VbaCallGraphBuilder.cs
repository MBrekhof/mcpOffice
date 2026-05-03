using System.Text.RegularExpressions;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static partial class VbaCallGraphBuilder
{
    // Reserved words / control-flow tokens that look like a procedure name but aren't a call.
    private static readonly HashSet<string> Keywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "If", "Then", "Else", "ElseIf", "End", "Sub", "Function", "Property",
        "Dim", "Set", "Const", "Public", "Private", "Friend", "Static",
        "For", "Next", "While", "Wend", "Do", "Loop", "Until", "Each", "In",
        "Select", "Case", "With", "Exit", "GoTo", "On", "Error", "Resume",
        "True", "False", "Nothing", "Null", "Empty", "And", "Or", "Not", "Xor", "Mod",
        "Is", "Like", "Optional", "ByVal", "ByRef", "ParamArray", "As",
        "Return", "Stop", "Rem", "Type", "Enum", "Declare", "Lib", "Alias",
        "Me", "Application"  // Application.Run is handled separately
    };

    [GeneratedRegex(@"^\s*(Call\s+)?(?<name>[A-Za-z_]\w*)\s*(\(|$)", RegexOptions.IgnoreCase)]
    private static partial Regex DirectCallRegex();

    [GeneratedRegex(@"Application\s*\.\s*Run\s+""(?<target>[^""]+)""", RegexOptions.IgnoreCase)]
    private static partial Regex AppRunRegex();

    [GeneratedRegex(@"^\s*[A-Za-z_]\w*\s*=", RegexOptions.IgnoreCase)]
    private static partial Regex AssignmentRegex();

    public static IReadOnlyList<ExcelVbaCallEdge> Build(
        IReadOnlyList<(string ModuleName, IReadOnlyList<CleanedLine> Lines, IReadOnlyList<ScannedProcedure> Procs)> modules)
    {
        // Build the procedure index: bare-name + FQN.
        var byBareName = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        var allFqn = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (_, _, procs) in modules)
        {
            foreach (var sp in procs)
            {
                allFqn.Add(sp.Procedure.FullyQualifiedName);
                if (!byBareName.TryGetValue(sp.Procedure.Name, out var list))
                {
                    list = [];
                    byBareName[sp.Procedure.Name] = list;
                }
                list.Add(sp.Procedure.FullyQualifiedName);
            }
        }

        var edges = new List<ExcelVbaCallEdge>();
        foreach (var (moduleName, lines, procs) in modules)
        {
            foreach (var sp in procs)
            {
                for (int i = sp.CleanedLineStartIndex; i <= sp.CleanedLineEndIndex && i < lines.Count; i++)
                {
                    var line = lines[i];

                    // Application.Run "Foo" — match against OriginalText because VbaLineCleaner
                    // replaces string literals with <STR> sentinels in the cleaned Text.
                    foreach (Match m in AppRunRegex().Matches(line.OriginalText))
                    {
                        edges.Add(new ExcelVbaCallEdge(
                            From: sp.Procedure.FullyQualifiedName,
                            To: m.Groups["target"].Value,
                            Resolved: false,
                            Site: new ExcelVbaSiteRef(moduleName, sp.Procedure.Name, line.LineNumber)));
                    }

                    // Skip lines that are obviously assignments
                    if (AssignmentRegex().IsMatch(line.Text)) continue;

                    var dc = DirectCallRegex().Match(line.Text);
                    if (!dc.Success) continue;

                    var name = dc.Groups["name"].Value;
                    if (Keywords.Contains(name)) continue;
                    // Skip self-name (the procedure header line was already excluded by [start, end] body window)
                    if (string.Equals(name, sp.Procedure.Name, StringComparison.OrdinalIgnoreCase)) continue;

                    string to;
                    bool resolved;
                    if (byBareName.TryGetValue(name, out var fqns))
                    {
                        // Prefer same-module match; otherwise first match
                        var sameMod = fqns.FirstOrDefault(f => f.StartsWith(moduleName + ".", StringComparison.OrdinalIgnoreCase));
                        to = sameMod ?? fqns[0];
                        resolved = true;
                    }
                    else
                    {
                        to = name;
                        resolved = false;
                    }

                    edges.Add(new ExcelVbaCallEdge(
                        From: sp.Procedure.FullyQualifiedName,
                        To: to,
                        Resolved: resolved,
                        Site: new ExcelVbaSiteRef(moduleName, sp.Procedure.Name, line.LineNumber)));
                }
            }
        }
        return edges;
    }
}
