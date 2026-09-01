using System.IO.Compression;
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel;

public class OpenXmlPartsTests
{
    private const string Rels = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Fact]
    public void ListSheets_follows_workbook_order_resolves_relative_and_absolute_targets_and_reads_codenames()
    {
        // rels deliberately list rId2 first; workbook.xml order (Blad1, Data) must win.
        using var zip = new ZipArchive(TestOpenXmlPackages.Build(
            ("xl/workbook.xml",
                $"""<workbook xmlns="{Main}" xmlns:r="{R}"><sheets><sheet name="Blad1" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>"""),
            ("xl/_rels/workbook.xml.rels",
                $"""<Relationships xmlns="{Rels}"><Relationship Id="rId2" Type="x/worksheet" Target="/xl/worksheets/sheet2.xml"/><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{Main}"><sheetPr codeName="Blad1"/><sheetData/></worksheet>"""),
            ("xl/worksheets/sheet2.xml", $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""")));

        var sheets = OpenXmlParts.ListSheets(zip);

        Assert.Equal(2, sheets.Count);
        Assert.Equal("Blad1", sheets[0].Name);
        Assert.Equal("Blad1", sheets[0].CodeName);
        Assert.Equal("xl/worksheets/sheet1.xml", sheets[0].PartPath);
        Assert.Equal("Data", sheets[1].Name);
        Assert.Null(sheets[1].CodeName);
        Assert.Equal("xl/worksheets/sheet2.xml", sheets[1].PartPath);
        Assert.Null(sheets[0].DrawingPartPath);
        Assert.Null(sheets[0].LegacyDrawingPartPath);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public void ListSheets_maps_drawing_and_legacy_drawing_parts(bool hasDrawing, bool hasLegacy)
    {
        using var zip = new ZipArchive(TestOpenXmlPackages.BuildWorkbookWithDrawings(
            drawingXml: hasDrawing ? "<xdr:wsDr xmlns:xdr=\"x\"/>" : null,
            vmlDrawing: hasLegacy ? "<xml/>" : null));

        var sheet = Assert.Single(OpenXmlParts.ListSheets(zip));

        Assert.Equal(hasDrawing ? "xl/drawings/drawing1.xml" : null, sheet.DrawingPartPath);
        Assert.Equal(hasLegacy ? "xl/drawings/vmlDrawing1.vml" : null, sheet.LegacyDrawingPartPath);
    }

    [Fact]
    public void ListSheets_returns_empty_without_workbook_part()
    {
        using var zip = new ZipArchive(TestOpenXmlPackages.Build(("docProps/app.xml", "<Properties/>")));
        Assert.Empty(OpenXmlParts.ListSheets(zip));
    }

    [Fact]
    public void ReadEntryText_returns_content_or_null()
    {
        using var zip = new ZipArchive(TestOpenXmlPackages.Build(("xl/drawings/vmlDrawing1.vml", "<xml>héllo</xml>")));

        Assert.Equal("<xml>héllo</xml>", OpenXmlParts.ReadEntryText(zip, "xl/drawings/vmlDrawing1.vml"));
        Assert.Null(OpenXmlParts.ReadEntryText(zip, "xl/drawings/missing.vml"));
    }

    [Fact]
    public void ReadFormulas_returns_non_empty_formula_texts_with_cells()
    {
        using var zip = new ZipArchive(TestOpenXmlPackages.BuildWorkbookWithDrawings(sheetData:
            $"""
            <sheetData>
              <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(A1:A5)</f><v>5</v></c></row>
              <row r="2"><c r="B2"><f t="shared" ref="B2:B3" si="0">A2*2</f><v>4</v></c></row>
              <row r="3"><c r="B3"><f t="shared" si="0"/><v>6</v></c><c r="C3"><f>MyUdf(A3)</f></c></row>
            </sheetData>
            """));

        var formulas = OpenXmlParts.ReadFormulas(zip, "xl/worksheets/sheet1.xml");

        Assert.Equal([("B1", "SUM(A1:A5)"), ("B2", "A2*2"), ("C3", "MyUdf(A3)")], formulas);
    }

    [Fact]
    public void ReadFormulas_returns_empty_for_missing_part()
    {
        using var zip = new ZipArchive(TestOpenXmlPackages.Build(("xl/workbook.xml", "<workbook/>")));
        Assert.Empty(OpenXmlParts.ReadFormulas(zip, "xl/worksheets/sheet9.xml"));
    }
}
