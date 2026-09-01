using System.Net;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace McpOffice.Services.Excel.Vba;

/// <summary>
/// Macro assignments on sheet shapes (`xl/drawings/drawingN.xml`, DrawingML) and legacy form
/// controls (`xl/drawings/vmlDrawingN.vml`). Pure string-in, records-out.
/// </summary>
internal static partial class DrawingMacroExtractor
{
    public sealed record ShapeMacro(string ShapeName, string ShapeKind, string MacroRef, string? TargetModule, string TargetProcedure, bool Parsable);

    private static readonly Dictionary<string, string> KindByElement = new(StringComparer.Ordinal)
    {
        ["sp"] = "shape",
        ["pic"] = "picture",
        ["cxnSp"] = "connector",
        ["graphicFrame"] = "graphicFrame",
    };

    [GeneratedRegex(@"^\[\d+\]!")]
    private static partial Regex ExternalIndexPrefixRegex();

    [GeneratedRegex(@"^[A-Za-z_][A-Za-z0-9_]*$")]
    private static partial Regex IdentifierRegex();

    // Excel lets a button run a macro with arguments: 'Copy_results(2)', 'Inlezen("Kjeldahl-N")'.
    // The procedure is the identifier before the argument list.
    [GeneratedRegex(@"^(?<name>[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?)\s*\(.*\)$", RegexOptions.Singleline)]
    private static partial Regex CallWithArgsRegex();

    // Lenient VML fallback: Excel's VML is frequently not well-formed XML.
    [GeneratedRegex(@"<(?:\w+:)?shape\b(?<attrs>[^>]*)>(?<body>.*?)</(?:\w+:)?shape>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex VmlShapeRegex();

    [GeneratedRegex(@"\bid\s*=\s*[""'](?<id>[^""']*)[""']", RegexOptions.IgnoreCase)]
    private static partial Regex VmlIdRegex();

    [GeneratedRegex(@"<(?:\w+:)?ClientData\b[^>]*\bObjectType\s*=\s*[""'](?<type>[^""']*)[""']", RegexOptions.IgnoreCase)]
    private static partial Regex VmlObjectTypeRegex();

    [GeneratedRegex(@"<(?:\w+:)?FmlaMacro>(?<macro>[^<]*)</(?:\w+:)?FmlaMacro>", RegexOptions.IgnoreCase)]
    private static partial Regex VmlFmlaMacroRegex();

    /// <summary>Throws <see cref="XmlException"/> on malformed input so the caller can count the skipped part.</summary>
    public static IReadOnlyList<ShapeMacro> FromDrawingXml(string xml)
    {
        var result = new List<ShapeMacro>();
        foreach (var element in XDocument.Parse(xml).Descendants().Where(e => KindByElement.ContainsKey(e.Name.LocalName)))
        {
            var macro = element.Attribute("macro")?.Value;
            if (string.IsNullOrWhiteSpace(macro)) continue;
            var name = element.Descendants().FirstOrDefault(d => d.Name.LocalName == "cNvPr")?.Attribute("name")?.Value ?? "";
            result.Add(Create(name, KindByElement[element.Name.LocalName], macro));
        }
        return result;
    }

    public static IReadOnlyList<ShapeMacro> FromVmlDrawing(string vml)
    {
        try
        {
            return FromVmlXml(XDocument.Parse(vml));
        }
        catch (XmlException)
        {
            return FromVmlRegex(vml);
        }
    }

    /// <summary>
    /// `[n]!` prefix stripped; `'Book'!Proc` is cross-workbook → unparsable; surrounding quotes
    /// stripped; a trailing argument list `(…)` dropped; then `Module.Proc` / `Proc` when both
    /// halves are VBA identifiers.
    /// </summary>
    public static (string? Module, string Procedure, bool Parsable) ParseMacroRef(string raw)
    {
        var s = ExternalIndexPrefixRegex().Replace(raw.Trim(), "");

        var bang = s.IndexOf("'!", StringComparison.Ordinal);
        if (s.StartsWith('\'') && bang > 0)
            return (null, s[(bang + 2)..].Trim(), false);

        if (s.Length >= 2 && s[0] == '\'' && s[^1] == '\'') s = s[1..^1].Trim();

        var call = CallWithArgsRegex().Match(s);
        if (call.Success) s = call.Groups["name"].Value;

        var dot = s.IndexOf('.');
        if (dot > 0 && IdentifierRegex().IsMatch(s[..dot]) && IdentifierRegex().IsMatch(s[(dot + 1)..]))
            return (s[..dot], s[(dot + 1)..], true);

        return (null, s, IdentifierRegex().IsMatch(s));
    }

    private static IReadOnlyList<ShapeMacro> FromVmlXml(XDocument doc)
    {
        var result = new List<ShapeMacro>();
        foreach (var shape in doc.Descendants().Where(e => e.Name.LocalName == "shape"))
        {
            var clientData = shape.Descendants().FirstOrDefault(e => e.Name.LocalName == "ClientData");
            AddVml(result,
                shape.Attribute("id")?.Value ?? "",
                clientData?.Attribute("ObjectType")?.Value,
                clientData?.Descendants().FirstOrDefault(e => e.Name.LocalName == "FmlaMacro")?.Value);
        }
        return result;
    }

    private static IReadOnlyList<ShapeMacro> FromVmlRegex(string vml)
    {
        var result = new List<ShapeMacro>();
        foreach (Match shape in VmlShapeRegex().Matches(vml))
        {
            var body = shape.Groups["body"].Value;
            var type = VmlObjectTypeRegex().Match(body);
            var macro = VmlFmlaMacroRegex().Match(body);
            AddVml(result,
                VmlIdRegex().Match(shape.Groups["attrs"].Value).Groups["id"].Value,
                type.Success ? type.Groups["type"].Value : null,
                macro.Success ? WebUtility.HtmlDecode(macro.Groups["macro"].Value) : null);   // &quot; etc., as XDocument would
        }
        return result;
    }

    private static void AddVml(List<ShapeMacro> sink, string id, string? objectType, string? macro)
    {
        if (string.IsNullOrWhiteSpace(objectType) || string.IsNullOrWhiteSpace(macro)) return;
        if (objectType.Equals("Note", StringComparison.OrdinalIgnoreCase)) return;   // cell comments
        sink.Add(Create(id, $"formControl:{objectType}", macro.Trim()));
    }

    private static ShapeMacro Create(string name, string kind, string macro)
    {
        var (module, proc, parsable) = ParseMacroRef(macro);
        return new ShapeMacro(name, kind, macro, module, proc, parsable);
    }
}
