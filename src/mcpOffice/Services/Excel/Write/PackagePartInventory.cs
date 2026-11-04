using System.IO.Compression;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Services.Excel.Write;

/// <summary>
/// Counts the package parts DevExpress is known to drop on a load/save round trip, so
/// <c>excel_update_workbook</c> can save anyway and warn (design doc, "Losing parts on update").
/// Compared by kind and count, never by bytes: DevExpress rewrites every part it keeps.
/// </summary>
internal static class PackagePartInventory
{
    public const string VbaProject = "vbaProject";
    public const string MacroButtonsPrefix = "macroButtons:";

    private static readonly (string Kind, Func<string, bool> Match, string Label)[] PartKinds =
    [
        ("activeX", p => p.StartsWith("xl/activeX/", StringComparison.OrdinalIgnoreCase) && p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase), "ActiveX controls"),
        ("slicers", p => p.StartsWith("xl/slicers/", StringComparison.OrdinalIgnoreCase), "slicers"),
        ("connections", p => string.Equals(p, "xl/connections.xml", StringComparison.OrdinalIgnoreCase), "external data connections"),
        ("queryTables", p => p.StartsWith("xl/queryTables/", StringComparison.OrdinalIgnoreCase), "query tables"),
        ("charts", p => p.StartsWith("xl/charts/chart", StringComparison.OrdinalIgnoreCase) && p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase), "charts"),
        ("pivotTables", p => p.StartsWith("xl/pivotTables/", StringComparison.OrdinalIgnoreCase), "pivot tables"),
    ];

    public static IReadOnlyDictionary<string, int> Take(ZipArchive zip)
    {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries)
        {
            if (string.Equals(entry.FullName, "xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase)) counts[VbaProject] = 1;
            foreach (var (kind, match, _) in PartKinds)
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
            if (buttons > 0) counts[MacroButtonsPrefix + sheet.Name] = buttons;
        }
        return counts;
    }

    /// <summary>One plain-language warning per kind whose count dropped.</summary>
    public static IReadOnlyList<string> Compare(IReadOnlyDictionary<string, int> before, IReadOnlyDictionary<string, int> after)
    {
        var warnings = new List<string>();
        foreach (var (kind, count) in before.OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            var lost = count - after.GetValueOrDefault(kind);
            if (lost <= 0) continue;
            if (kind == VbaProject)
                warnings.Add("The VBA project was not preserved");
            else if (kind.StartsWith(MacroButtonsPrefix, StringComparison.Ordinal))
                warnings.Add($"{lost} form-control button{(lost == 1 ? "" : "s")} with macro assignments on sheet '{kind[MacroButtonsPrefix.Length..]}' {(lost == 1 ? "was" : "were")} not preserved");
            else
                warnings.Add($"{lost} {PartKinds.First(k => k.Kind == kind).Label} {(lost == 1 ? "was" : "were")} not preserved");
        }
        return warnings;
    }
}
