using System.Xml;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class DrawingMacroExtractorTests
{
    [Theory]
    [InlineData("[0]!GetILIS", null, "GetILIS", true)]
    [InlineData("[0]!Inlezen", null, "Inlezen", true)]
    [InlineData("'Copy_results(2)'", null, "Copy_results", true)]
    [InlineData("Module1.Foo", "Module1", "Foo", true)]
    [InlineData("[12]!Module1.Foo", "Module1", "Foo", true)]
    [InlineData("'Book.xlsm'!Proc", null, "Proc", false)]
    [InlineData("'Book.xlsm'!Module1.Proc", null, "Module1.Proc", false)]
    [InlineData("  Proc  ", null, "Proc", true)]
    [InlineData("'Proc'", null, "Proc", true)]
    [InlineData("Sheet1.Foo.Bar", null, "Sheet1.Foo.Bar", false)]
    [InlineData("1Bad", null, "1Bad", false)]
    public void ParseMacroRef_follows_the_rules(string raw, string? module, string procedure, bool parsable)
    {
        Assert.Equal((module, procedure, parsable), DrawingMacroExtractor.ParseMacroRef(raw));
    }

    private const string DrawingXml = """
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <xdr:twoCellAnchor>
            <xdr:sp macro="[0]!GetILIS" textlink="">
              <xdr:nvSpPr><xdr:cNvPr id="15" name="Button 14"/><xdr:cNvSpPr/></xdr:nvSpPr>
              <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
            </xdr:sp>
          </xdr:twoCellAnchor>
          <xdr:twoCellAnchor editAs="oneCell">
            <xdr:pic macro="">
              <xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 2"/><xdr:cNvPicPr/></xdr:nvPicPr>
              <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
            </xdr:pic>
          </xdr:twoCellAnchor>
          <xdr:twoCellAnchor>
            <xdr:grpSp>
              <xdr:nvGrpSpPr><xdr:cNvPr id="20" name="Group 19"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
              <xdr:pic macro="Module1.Inlezen">
                <xdr:nvPicPr><xdr:cNvPr id="85" name="Picture 84"/><xdr:cNvPicPr/></xdr:nvPicPr>
              </xdr:pic>
              <xdr:cxnSp macro="'Copy_results(2)'">
                <xdr:nvCxnSpPr><xdr:cNvPr id="21" name="Straight Connector 20"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>
              </xdr:cxnSp>
            </xdr:grpSp>
          </xdr:twoCellAnchor>
          <xdr:twoCellAnchor>
            <xdr:graphicFrame macro="[0]!Chart_Click">
              <xdr:nvGraphicFramePr><xdr:cNvPr id="7" name="Chart 6"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
              <a:graphic><a:graphicData uri="x"/></a:graphic>
            </xdr:graphicFrame>
          </xdr:twoCellAnchor>
        </xdr:wsDr>
        """;

    [Fact]
    public void FromDrawingXml_extracts_shapes_pictures_connectors_and_frames_including_grouped_and_skips_empty_macro()
    {
        var macros = DrawingMacroExtractor.FromDrawingXml(DrawingXml);

        Assert.Equal(
        [
            new DrawingMacroExtractor.ShapeMacro("Button 14", "shape", "[0]!GetILIS", null, "GetILIS", true),
            new DrawingMacroExtractor.ShapeMacro("Picture 84", "picture", "Module1.Inlezen", "Module1", "Inlezen", true),
            new DrawingMacroExtractor.ShapeMacro("Straight Connector 20", "connector", "'Copy_results(2)'", null, "Copy_results", true),
            new DrawingMacroExtractor.ShapeMacro("Chart 6", "graphicFrame", "[0]!Chart_Click", null, "Chart_Click", true),
        ], macros);
    }

    [Fact]
    public void FromDrawingXml_throws_XmlException_on_malformed_input_so_the_caller_can_count_the_skip()
    {
        Assert.Throws<XmlException>(() => DrawingMacroExtractor.FromDrawingXml("<xdr:wsDr><xdr:sp macro=\"a\">"));
    }

    private const string Vml = """
        <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
         <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
         <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">
          <v:stroke joinstyle="miter"/>
         </v:shapetype>
         <v:shape id="_x0000_s1025" type="#_x0000_t201" style='position:absolute;margin-left:1pt' o:button="t" o:insetmode="auto">
          <v:textbox style='mso-direction-alt:auto' o:singleclick="f"><div style='text-align:center'>Next</div></v:textbox>
          <x:ClientData ObjectType="Button">
           <x:Anchor>1, 0, 0, 0, 2, 0, 1, 0</x:Anchor>
           <x:PrintObject>False</x:PrintObject>
           <x:FmlaMacro>[0]!NextDate</x:FmlaMacro>
           <x:TextHAlign>Center</x:TextHAlign>
          </x:ClientData>
         </v:shape>
         <v:shape id="_x0000_s1026" type="#_x0000_t202" style='position:absolute'>
          <v:textbox><div style='text-align:left'>A comment</div></v:textbox>
          <x:ClientData ObjectType="Note"><x:MoveWithCells/><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData>
         </v:shape>
         <v:shape id="_x0000_s1027" type="#_x0000_t201" style='position:absolute'>
          <x:ClientData ObjectType="Checkbox">
           <x:FmlaMacro>Module1.Toggle</x:FmlaMacro>
           <x:Checked>1</x:Checked>
          </x:ClientData>
         </v:shape>
         <v:shape id="_x0000_s1028" type="#_x0000_t201" style='position:absolute'>
          <x:ClientData ObjectType="Drop"><x:DropLines>8</x:DropLines></x:ClientData>
         </v:shape>
         <v:shape id="_x0000_s1029" type="#_x0000_t201" style='position:absolute'>
          <x:ClientData ObjectType="Button"><x:FmlaMacro>'ApparaatInlezen(&quot;Kjeldahl-N&quot;)'</x:FmlaMacro></x:ClientData>
         </v:shape>
        </xml>
        """;

    private static readonly DrawingMacroExtractor.ShapeMacro[] ExpectedVml =
    [
        new("_x0000_s1025", "formControl:Button", "[0]!NextDate", null, "NextDate", true),
        new("_x0000_s1027", "formControl:Checkbox", "Module1.Toggle", "Module1", "Toggle", true),
        // Air.xlsm shape: a quoted call expression with XML-escaped quotes — decoded, kept verbatim, unparsable.
        new("_x0000_s1029", "formControl:Button", "'ApparaatInlezen(\"Kjeldahl-N\")'", null, "ApparaatInlezen", true),
    ];

    [Fact]
    public void FromVmlDrawing_well_formed_skips_notes_and_controls_without_macro()
    {
        Assert.Equal(ExpectedVml, DrawingMacroExtractor.FromVmlDrawing(Vml));
    }

    [Fact]
    public void FromVmlDrawing_falls_back_to_regex_when_not_well_formed()
    {
        // Unclosed <br> and a conditional-comment construct: XDocument.Parse rejects both.
        var broken = Vml.Replace("<div style='text-align:center'>Next</div>", "<![if !vml]><div>Next<br>Date</div><![endif]>");
        Assert.Throws<XmlException>(() => System.Xml.Linq.XDocument.Parse(broken));

        Assert.Equal(ExpectedVml, DrawingMacroExtractor.FromVmlDrawing(broken));
    }
}
