using System.IO.Compression;
using McpOffice.Services.Excel.Write;

namespace McpOffice.Tests.Excel.Write;

public class PackagePartInventoryTests
{
    private const string DrawingWithTwoButtons = """
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xdr:twoCellAnchor><xdr:sp macro="[0]!Inlezen"><xdr:nvSpPr><xdr:cNvPr id="1" name="Button 1"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor>
          <xdr:twoCellAnchor><xdr:sp macro="[0]!Export"><xdr:nvSpPr><xdr:cNvPr id="2" name="Button 2"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor>
          <xdr:twoCellAnchor><xdr:sp macro=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Plain shape"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor>
        </xdr:wsDr>
        """;

    private const string VmlWithOneButton = """
        <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
          <v:shape id="_x0000_s1025"><x:ClientData ObjectType="Button"><x:FmlaMacro>[0]!Knop_Click</x:FmlaMacro></x:ClientData></v:shape>
          <v:shape id="_x0000_s1026"><x:ClientData ObjectType="Note"><x:FmlaMacro>ignored</x:FmlaMacro></x:ClientData></v:shape>
        </xml>
        """;

    [Fact]
    public void Take_counts_each_kind()
    {
        using var ms = TestOpenXmlPackages.BuildWorkbookWithDrawings(sheetName: "Invoer", drawingXml: DrawingWithTwoButtons, vmlDrawing: VmlWithOneButton);
        // Append the other kinds to the same package.
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var part in new[] { "xl/vbaProject.bin", "xl/activeX/activeX1.xml", "xl/activeX/activeX1.bin", "xl/charts/chart1.xml", "xl/charts/chart2.xml", "xl/charts/style1.xml", "xl/slicers/slicer1.xml", "xl/connections.xml", "xl/queryTables/queryTable1.xml", "xl/pivotTables/pivotTable1.xml" })
                zip.CreateEntry(part);
        }
        ms.Position = 0;

        using var read = new ZipArchive(ms, ZipArchiveMode.Read);
        var counts = PackagePartInventory.Take(read);

        Assert.Equal(1, counts["vbaProject"]);
        Assert.Equal(3, counts["macroButtons:Invoer"]);
        Assert.Equal(1, counts["activeX"]);
        Assert.Equal(2, counts["charts"]);
        Assert.Equal(1, counts["slicers"]);
        Assert.Equal(1, counts["connections"]);
        Assert.Equal(1, counts["queryTables"]);
        Assert.Equal(1, counts["pivotTables"]);
    }

    [Fact]
    public void Take_of_a_plain_workbook_is_empty()
    {
        using var ms = TestOpenXmlPackages.BuildWorkbookWithDrawings();
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.Empty(PackagePartInventory.Take(zip));
    }

    [Fact]
    public void Compare_equal_maps_gives_no_warnings()
    {
        var map = new Dictionary<string, int> { ["vbaProject"] = 1, ["macroButtons:Invoer"] = 3, ["charts"] = 2 };
        Assert.Empty(PackagePartInventory.Compare(map, map));
        Assert.Empty(PackagePartInventory.Compare(new Dictionary<string, int>(), map));   // gained parts are not a loss
    }

    [Fact]
    public void Compare_reports_each_drop_in_plain_language()
    {
        var before = new Dictionary<string, int> { ["vbaProject"] = 1, ["macroButtons:Invoer"] = 3, ["charts"] = 2, ["activeX"] = 1 };
        var after = new Dictionary<string, int> { ["charts"] = 1, ["macroButtons:Invoer"] = 1 };

        Assert.Equal(
        [
            "1 ActiveX controls was not preserved",
            "1 charts was not preserved",
            "2 form-control buttons with macro assignments on sheet 'Invoer' were not preserved",
            "The VBA project was not preserved",
        ], PackagePartInventory.Compare(before, after));
    }
}
