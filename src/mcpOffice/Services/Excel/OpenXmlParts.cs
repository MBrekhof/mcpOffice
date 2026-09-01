using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace McpOffice.Services.Excel;

/// <summary>
/// Pure functions over an OOXML package: sheet name ↔ part ↔ codename ↔ drawing parts, plus raw
/// part text and cell formulas. Namespace-agnostic element matching (LocalName) throughout.
/// </summary>
internal static class OpenXmlParts
{
    public sealed record SheetPart(string Name, string? CodeName, string PartPath, string? DrawingPartPath, string? LegacyDrawingPartPath);

    private static readonly XmlReaderSettings ReaderSettings = new()
    {
        DtdProcessing = DtdProcessing.Ignore,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
    };

    /// <summary>Sheets in workbook order. Sheets whose relationship cannot be resolved are skipped.</summary>
    public static IReadOnlyList<SheetPart> ListSheets(ZipArchive zip)
    {
        var workbook = LoadXml(zip, "xl/workbook.xml");
        if (workbook is null) return [];

        var workbookRels = ReadRelationships(zip, "xl/workbook.xml");
        var result = new List<SheetPart>();
        foreach (var sheet in workbook.Descendants().Where(e => e.Name.LocalName == "sheet"))
        {
            var rid = sheet.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value;
            if (rid is null || !workbookRels.TryGetValue(rid, out var partPath)) continue;

            var sheetXml = LoadXml(zip, partPath);
            var sheetRels = ReadRelationships(zip, partPath);
            result.Add(new SheetPart(
                sheet.Attribute("name")?.Value ?? "",
                sheetXml?.Descendants().FirstOrDefault(e => e.Name.LocalName == "sheetPr")?.Attribute("codeName")?.Value,
                partPath,
                ResolveReference(sheetXml, "drawing", sheetRels),
                ResolveReference(sheetXml, "legacyDrawing", sheetRels)));
        }
        return result;
    }

    /// <summary>UTF-8 text of a part, or null when the part is missing.</summary>
    public static string? ReadEntryText(ZipArchive zip, string partPath)
    {
        var entry = FindEntry(zip, partPath);
        if (entry is null) return null;
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    /// <summary>
    /// Every non-empty &lt;f&gt; in a sheet part with its cell address. Shared-formula children
    /// (`&lt;f t="shared" si="0"/&gt;`) carry no text and are skipped; the master's text suffices.
    /// </summary>
    public static IReadOnlyList<(string Cell, string Formula)> ReadFormulas(ZipArchive zip, string sheetPartPath)
    {
        var entry = FindEntry(zip, sheetPartPath);
        if (entry is null) return [];

        var result = new List<(string, string)>();
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, ReaderSettings);
        var cell = "";
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.LocalName == "c") cell = reader.GetAttribute("r") ?? "";
            else if (reader.LocalName == "f" && !reader.IsEmptyElement)
            {
                var formula = reader.ReadElementContentAsString();
                if (!string.IsNullOrWhiteSpace(formula)) result.Add((cell, formula));
            }
        }
        return result;
    }

    private static string? ResolveReference(XDocument? sheetXml, string elementLocalName, Dictionary<string, string> rels)
    {
        var rid = sheetXml?.Descendants().FirstOrDefault(e => e.Name.LocalName == elementLocalName)
            ?.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value;
        return rid is not null && rels.TryGetValue(rid, out var target) ? target : null;
    }

    /// <summary>Id → resolved part path for the .rels of <paramref name="partPath"/>. Empty when absent.</summary>
    private static Dictionary<string, string> ReadRelationships(ZipArchive zip, string partPath)
    {
        var folder = FolderOf(partPath);
        var relsPath = $"{folder}/_rels/{partPath[(folder.Length + 1)..]}.rels";
        var rels = new Dictionary<string, string>(StringComparer.Ordinal);
        var doc = LoadXml(zip, relsPath);
        if (doc is null) return rels;

        foreach (var rel in doc.Descendants().Where(e => e.Name.LocalName == "Relationship"))
        {
            var id = rel.Attribute("Id")?.Value;
            var target = rel.Attribute("Target")?.Value;
            if (id is not null && target is not null) rels[id] = ResolveTarget(folder, target);
        }
        return rels;
    }

    private static string ResolveTarget(string baseFolder, string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var segments = new List<string>(baseFolder.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (var segment in target.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); }
            else if (segment != ".") segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    private static string FolderOf(string partPath)
    {
        var slash = partPath.LastIndexOf('/');
        return slash < 0 ? "" : partPath[..slash];
    }

    private static XDocument? LoadXml(ZipArchive zip, string partPath)
    {
        var entry = FindEntry(zip, partPath);
        if (entry is null) return null;
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, ReaderSettings);
        return XDocument.Load(reader);
    }

    // OPC part names are case-insensitive; ZipArchive.GetEntry is not.
    private static ZipArchiveEntry? FindEntry(ZipArchive zip, string partPath) =>
        zip.GetEntry(partPath)
        ?? zip.Entries.FirstOrDefault(e => string.Equals(e.FullName, partPath, StringComparison.OrdinalIgnoreCase));
}
