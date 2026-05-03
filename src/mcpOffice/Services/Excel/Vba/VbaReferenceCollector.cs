using System.Text.RegularExpressions;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static partial class VbaReferenceCollector
{
    // Present for future enum extraction; the regex hardcodes the alternation.
    private static readonly string[] ObjectModelApis =
    [
        "Worksheets", "Sheets", "Range", "Cells",
        "ActiveSheet", "ActiveWorkbook", "ThisWorkbook",
        "Application", "Selection", "Names"
    ];

    // First pass: APIs invoked with a string-literal arg — capture the literal from OriginalText.
    [GeneratedRegex(@"\b(?<api>Worksheets|Sheets|Range|Cells|ActiveSheet|ActiveWorkbook|ThisWorkbook|Application|Selection|Names)\b\s*\(\s*""(?<lit>[^""]*)""", RegexOptions.IgnoreCase)]
    private static partial Regex OmWithLiteralRegex();

    // Second pass: bare references that have no string-literal arg (or were not captured above).
    [GeneratedRegex(@"\b(?<api>Worksheets|Sheets|Range|Cells|ActiveSheet|ActiveWorkbook|ThisWorkbook|Application|Selection|Names)\b", RegexOptions.IgnoreCase)]
    private static partial Regex OmAnyRegex();

    [GeneratedRegex(@"^\s*Open\s+", RegexOptions.IgnoreCase)]
    private static partial Regex FileOpenRegex();

    [GeneratedRegex(@"\b(Kill|MkDir|RmDir|ChDir|Dir|FileSystemObject|Workbooks\.Open|Workbooks\.OpenText|Scripting\.FileSystemObject)\b", RegexOptions.IgnoreCase)]
    private static partial Regex FileApiRegex();

    [GeneratedRegex(@"(?:CreateObject|GetObject)\s*\(\s*""(?<progid>[^""]+)""", RegexOptions.IgnoreCase)]
    private static partial Regex CreateGetObjectRegex();

    [GeneratedRegex(@"\b(ADODB\.|DAO\.|OpenDatabase)\b", RegexOptions.IgnoreCase)]
    private static partial Regex DatabaseApiRegex();

    [GeneratedRegex(@"\b(MSXML2\.XMLHTTP|WinHttp\.WinHttpRequest|URLDownloadToFile|InternetExplorer\.Application)\b", RegexOptions.IgnoreCase)]
    private static partial Regex NetworkApiRegex();

    [GeneratedRegex(@"^\s*Shell\s*\(", RegexOptions.IgnoreCase)]
    private static partial Regex ShellRegex();

    public static void Collect(
        string moduleName,
        IReadOnlyList<CleanedLine> lines,
        IReadOnlyList<ScannedProcedure> procs,
        List<ExcelVbaObjectModelRef> objectModelOut,
        List<ExcelVbaDependency> dependenciesOut)
    {
        foreach (var sp in procs)
        {
            for (int i = sp.CleanedLineStartIndex; i <= sp.CleanedLineEndIndex && i < lines.Count; i++)
            {
                var line = lines[i];

                CollectObjectModel(moduleName, sp.Procedure.Name, line, objectModelOut);
                CollectDependencies(moduleName, sp.Procedure.Name, line, dependenciesOut);
            }
        }
    }

    private static void CollectObjectModel(string module, string proc, CleanedLine line, List<ExcelVbaObjectModelRef> sink)
    {
        var seenAt = new HashSet<int>();

        // First pass: APIs invoked with a string-literal arg — capture from OriginalText.
        foreach (Match m in OmWithLiteralRegex().Matches(line.OriginalText))
        {
            sink.Add(new ExcelVbaObjectModelRef(module, proc, line.LineNumber, m.Groups["api"].Value, m.Groups["lit"].Value));
            seenAt.Add(m.Index);
        }

        // Second pass: bare references (no literal arg). De-dupe by start position to avoid double-counting.
        foreach (Match m in OmAnyRegex().Matches(line.OriginalText))
        {
            if (seenAt.Contains(m.Index)) continue;
            sink.Add(new ExcelVbaObjectModelRef(module, proc, line.LineNumber, m.Groups["api"].Value, null));
        }
    }

    private static void CollectDependencies(string module, string proc, CleanedLine line, List<ExcelVbaDependency> sink)
    {
        // Shell builtin (statement form, leading)
        if (ShellRegex().IsMatch(line.Text))
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "shell", null, "Shell"));
            return;
        }

        // File: Open ... For ...
        if (FileOpenRegex().IsMatch(line.Text))
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "file", null, "Open"));
            return;
        }

        var fileM = FileApiRegex().Match(line.Text);
        if (fileM.Success)
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "file", null, fileM.Groups[1].Value));
            return;
        }

        // CreateObject / GetObject — dispatch by ProgID.
        var co = CreateGetObjectRegex().Match(line.OriginalText);
        if (co.Success)
        {
            var progId = co.Groups["progid"].Value;
            string kind = ClassifyProgId(progId);
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, kind, progId, "CreateObject"));
            return;
        }

        if (DatabaseApiRegex().IsMatch(line.Text))
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "database", null, null));
            return;
        }
        if (NetworkApiRegex().IsMatch(line.Text))
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "network", null, null));
            return;
        }
    }

    private static string ClassifyProgId(string progId)
    {
        if (progId.StartsWith("ADODB.", StringComparison.OrdinalIgnoreCase) ||
            progId.StartsWith("DAO.", StringComparison.OrdinalIgnoreCase)) return "database";
        if (progId.StartsWith("MSXML2.", StringComparison.OrdinalIgnoreCase) ||
            progId.StartsWith("WinHttp.", StringComparison.OrdinalIgnoreCase) ||
            progId.Equals("InternetExplorer.Application", StringComparison.OrdinalIgnoreCase)) return "network";
        if (progId.Contains("FileSystemObject", StringComparison.OrdinalIgnoreCase)) return "file";
        return "automation";
    }
}
