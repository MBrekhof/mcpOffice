using System.IO.Compression;
using System.Text.RegularExpressions;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Services.Excel.Write;

/// <summary>
/// Counts the package parts DevExpress is known to drop on a load/save round trip, so
/// <c>excel_update_workbook</c> can save anyway and warn (design doc, "Losing parts on update").
/// Compared by kind and count, never by bytes: DevExpress rewrites every part it keeps.
/// </summary>
internal static partial class PackagePartInventory
{
    public const string VbaProject = "vbaProject";
    public const string MacroButtonsPrefix = "macroButtons:";

    // Only the main part of each kind: chartStyle1.xml / colors1.xml and the _rels siblings are not
    // separate objects, and counting them would report one lost chart as three.
    [GeneratedRegex(@"^xl/charts/chart(Ex)?\d+\.xml$", RegexOptions.IgnoreCase)]
    private static partial Regex ChartPartRegex();

    private static readonly (string Kind, Func<string, bool> Match, string Singular, string Plural)[] PartKinds =
    [
        ("activeX", p => IsMainPartUnder("xl/activeX/", p), "ActiveX control", "ActiveX controls"),
        ("slicers", p => IsMainPartUnder("xl/slicers/", p), "slicer", "slicers"),
        ("connections", p => string.Equals(p, "xl/connections.xml", StringComparison.OrdinalIgnoreCase), "external data connections part", "external data connections parts"),
        ("queryTables", p => IsMainPartUnder("xl/queryTables/", p), "query table", "query tables"),
        ("charts", p => ChartPartRegex().IsMatch(p), "chart", "charts"),
        ("pivotTables", p => IsMainPartUnder("xl/pivotTables/", p), "pivot table", "pivot tables"),
    ];

    private static bool IsMainPartUnder(string folder, string path) =>
        path.StartsWith(folder, StringComparison.OrdinalIgnoreCase)
        && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
        && !path.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);

    public static IReadOnlyDictionary<string, int> Take(ZipArchive zip)
    {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries)
        {
            if (string.Equals(entry.FullName, "xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase)) counts[VbaProject] = 1;
            foreach (var (kind, match, _, _) in PartKinds)
                if (match(entry.FullName)) counts[kind] = counts.GetValueOrDefault(kind) + 1;
        }

        foreach (var sheet in OpenXmlParts.ListSheets(zip))
        {
            var buttons = 0;
            if (sheet.DrawingPartPath is not null && OpenXmlParts.ReadEntryText(zip, sheet.DrawingPartPath) is string drawing)
            {
                try { buttons += DrawingMacroExtractor.FromDrawingXml(drawing).Count; }
                catch (System.Xml.XmlException) { /* malformed part: nothing countable */ }
            }
            if (sheet.LegacyDrawingPartPath is not null && OpenXmlParts.ReadEntryText(zip, sheet.LegacyDrawingPartPath) is string vml)
                buttons += DrawingMacroExtractor.FromVmlDrawing(vml).Count;
            // Keyed by code name where there is one, so a renameSheet in the same call is not reported as a loss.
            if (buttons > 0) counts[MacroButtonsPrefix + (sheet.CodeName ?? sheet.Name)] = buttons;
        }
        return counts;
    }

    /// <summary>One plain-language warning per kind whose count dropped.</summary>
    /// <param name="sheetName">Maps the sheet id in a <c>macroButtons:</c> key (code name) to the name the caller knows.</param>
    public static IReadOnlyList<string> Compare(IReadOnlyDictionary<string, int> before, IReadOnlyDictionary<string, int> after, Func<string, string>? sheetName = null)
    {
        var warnings = new List<string>();
        foreach (var (kind, count) in before.OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            var lost = count - after.GetValueOrDefault(kind);
            if (lost <= 0) continue;
            if (kind == VbaProject)
                warnings.Add("The VBA project was not preserved");
            else if (kind.StartsWith(MacroButtonsPrefix, StringComparison.Ordinal))
            {
                var id = kind[MacroButtonsPrefix.Length..];
                var name = sheetName?.Invoke(id) ?? id;
                warnings.Add(lost == 1
                    ? $"1 form-control button with a macro assignment on sheet '{name}' was not preserved"
                    : $"{lost} form-control buttons with macro assignments on sheet '{name}' were not preserved");
            }
            else
            {
                var (_, _, singular, plural) = PartKinds.First(k => k.Kind == kind);
                warnings.Add(lost == 1 ? $"1 {singular} was not preserved" : $"{lost} {plural} were not preserved");
            }
        }
        return warnings;
    }
}
