using System.IO.Compression;
using System.Text;

namespace McpOffice.Tests.Excel;

/// <summary>
/// Builds minimal OOXML packages in memory. No [Content_Types].xml — nothing under test validates it.
/// </summary>
internal static class TestOpenXmlPackages
{
    public static MemoryStream Build(params (string Path, string Content)[] parts)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (path, content) in parts)
            {
                using var writer = new StreamWriter(zip.CreateEntry(path).Open(), new UTF8Encoding(false));
                writer.Write(content);
            }
        }
        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// One-sheet workbook. <paramref name="drawingXml"/> / <paramref name="vmlDrawing"/> null = no
    /// &lt;drawing&gt; / &lt;legacyDrawing&gt; element and no part.
    /// </summary>
    public static MemoryStream BuildWorkbookWithDrawings(
        string sheetName = "Blad1",
        string? codeName = "Blad1",
        string? drawingXml = null,
        string? vmlDrawing = null,
        string sheetData = "<sheetData/>")
    {
        var parts = new List<(string, string)>
        {
            ("xl/workbook.xml",
                $"""<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr codeName="ThisWorkbook"/><sheets><sheet name="{sheetName}" sheetId="1" r:id="rId1"/></sheets></workbook>"""),
            ("xl/_rels/workbook.xml.rels",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""),
        };

        var sheetPr = codeName is null ? "" : $"""<sheetPr codeName="{codeName}"/>""";
        var drawingEl = drawingXml is null ? "" : """<drawing r:id="rId1"/>""";
        var legacyEl = vmlDrawing is null ? "" : """<legacyDrawing r:id="rId2"/>""";
        parts.Add(("xl/worksheets/sheet1.xml",
            $"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{sheetPr}{sheetData}{drawingEl}{legacyEl}</worksheet>"""));

        var rels = new StringBuilder("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""");
        if (drawingXml is not null)
        {
            rels.Append("""<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>""");
            parts.Add(("xl/drawings/drawing1.xml", drawingXml));
        }
        if (vmlDrawing is not null)
        {
            rels.Append("""<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>""");
            parts.Add(("xl/drawings/vmlDrawing1.vml", vmlDrawing));
        }
        rels.Append("</Relationships>");
        parts.Add(("xl/worksheets/_rels/sheet1.xml.rels", rels.ToString()));

        return Build(parts.ToArray());
    }
}
