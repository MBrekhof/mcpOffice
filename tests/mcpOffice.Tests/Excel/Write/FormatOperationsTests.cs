using DevExpress.Spreadsheet;
using McpOffice.Models;
using ModelContextProtocol;
using static McpOffice.Tests.Excel.Write.WriteTestSupport;

namespace McpOffice.Tests.Excel.Write;

public class FormatOperationsTests
{
    private static string Input() => TestExcelWorkbooks.Create(wb =>
    {
        var ws = wb.Worksheets[0];
        ws.Name = "Data";
        for (var r = 0; r < 4; r++)
            for (var c = 0; c < 3; c++)
                ws.Cells[r, c].Value = r * 10 + c;
        ws.Range["A1:C1"].Font.Bold = true;
    });

    private static Workbook Apply(string input, string output, params ExcelOperation[] ops)
    {
        Update(input, output, ops);
        return Load(output);
    }

    [Fact]
    public void FormatRange_sets_only_given_properties()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            using var wb = Apply(input, output, new ExcelOperation("formatRange", Sheet: "Data", Range: "A1:C1", FillColor: "#FFFF00"));
            var a1 = wb.Worksheets[0].Cells["A1"];
            Assert.True(a1.Font.Bold);
            Assert.Equal(0xFFFF00, a1.Fill.BackgroundColor.ToArgb() & 0xFFFFFF);
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void FormatRange_font_fill_number_alignment_and_wrap()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            using var wb = Apply(input, output,
                new ExcelOperation("formatRange", Sheet: "Data", Range: "A2", Bold: true, Italic: true, Underline: true, FontColor: "#FF0000", FontSize: 14),
                new ExcelOperation("formatRange", Sheet: "Data", Range: "B2", NumberFormat: "0.00", HorizontalAlign: "center", VerticalAlign: "top", WrapText: true),
                new ExcelOperation("formatRange", Sheet: "Data", Range: "C2", HorizontalAlign: "right", VerticalAlign: "bottom"),
                new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", Bold: false));
            var ws = wb.Worksheets[0];
            var a2 = ws.Cells["A2"];
            Assert.True(a2.Font.Bold);
            Assert.True(a2.Font.Italic);
            Assert.Equal(UnderlineType.Single, a2.Font.UnderlineType);
            Assert.Equal(0xFF0000, a2.Font.Color.ToArgb() & 0xFFFFFF);
            Assert.Equal(14d, a2.Font.Size);
            var b2 = ws.Cells["B2"];
            Assert.Equal("0.00", b2.NumberFormat);
            Assert.Equal(SpreadsheetHorizontalAlignment.Center, b2.Alignment.Horizontal);
            Assert.Equal(SpreadsheetVerticalAlignment.Top, b2.Alignment.Vertical);
            Assert.True(b2.Alignment.WrapText);
            Assert.Equal(SpreadsheetHorizontalAlignment.Right, ws.Cells["C2"].Alignment.Horizontal);
            Assert.Equal(SpreadsheetVerticalAlignment.Bottom, ws.Cells["C2"].Alignment.Vertical);
            Assert.False(ws.Cells["A1"].Font.Bold);
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Border_all_versus_outline()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            using var wb = Apply(input, output,
                new ExcelOperation("formatRange", Sheet: "Data", Range: "A1:B2", Border: "thin", BorderScope: "outline"),
                new ExcelOperation("formatRange", Sheet: "Data", Range: "A3:B4", Border: "medium"));
            var ws = wb.Worksheets[0];
            Assert.Equal(BorderLineStyle.Thin, ws.Cells["A1"].Borders.TopBorder.LineStyle);
            Assert.Equal(BorderLineStyle.Thin, ws.Cells["A1"].Borders.LeftBorder.LineStyle);
            Assert.Equal(BorderLineStyle.None, ws.Cells["A1"].Borders.RightBorder.LineStyle);   // inner edge
            Assert.Equal(BorderLineStyle.Medium, ws.Cells["A3"].Borders.RightBorder.LineStyle);  // inner edge, scope all
            Assert.Equal(BorderLineStyle.Medium, ws.Cells["B4"].Borders.BottomBorder.LineStyle);
        }
        finally { Delete(input, output); }
    }

    [Theory]
    [InlineData("horizontalAlign", "middle")]
    [InlineData("verticalAlign", "center")]
    [InlineData("border", "dotted")]
    [InlineData("fillColor", "red")]
    [InlineData("fontColor", "#FFF")]
    [InlineData("borderScope", "inside")]
    [InlineData("range", "not a range")]
    public void Bad_values_are_invalid_operation_with_index(string field, string value)
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            var op = field switch
            {
                "horizontalAlign" => new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", HorizontalAlign: value),
                "verticalAlign" => new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", VerticalAlign: value),
                "border" => new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", Border: value),
                "fillColor" => new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", FillColor: value),
                "fontColor" => new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", FontColor: value),
                "borderScope" => new ExcelOperation("formatRange", Sheet: "Data", Range: "A1", Border: "thin", BorderScope: value),
                _ => new ExcelOperation("formatRange", Sheet: "Data", Range: value, Bold: true),
            };
            var ex = Assert.Throws<McpException>(() => Update(input, output, new ExcelOperation("addSheet", Name: "X"), op));
            Assert.StartsWith("[invalid_operation] operations[1] (formatRange): ", ex.Message);
            Assert.Contains(value, ex.Message);
            Assert.False(File.Exists(output));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void FormatRange_without_properties_is_invalid_operation()
    {
        var input = Input();
        try
        {
            var ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("formatRange", Sheet: "Data", Range: "A1")));
            Assert.StartsWith("[invalid_operation] operations[0] (formatRange): formatRange needs at least one", ex.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void MergeCells_and_unmergeCells_round_trip()
    {
        var input = Input();
        var merged = TempPath(".xlsx");
        var unmerged = TempPath(".xlsx");
        try
        {
            using (var wb = Apply(input, merged, new ExcelOperation("mergeCells", Sheet: "Data", Range: "A1:C1")))
                Assert.True(wb.Worksheets[0].Range["A1:C1"].IsMerged);
            using (var wb = Apply(merged, unmerged, new ExcelOperation("unmergeCells", Sheet: "Data", Range: "A1:C1")))
                Assert.Empty(wb.Worksheets[0].Range["A1:C1"].GetMergedRanges());
        }
        finally { Delete(input, merged, unmerged); }
    }

    [Fact]
    public void SetColumnWidth_span_autofit_and_validation()
    {
        var input = TestExcelWorkbooks.Create(wb =>
        {
            var ws = wb.Worksheets[0];
            ws.Name = "Data";
            ws.Cells["A1"].Value = "x";
            ws.Cells["D1"].Value = new string('y', 300);
        });
        var output = TempPath(".xlsx");
        try
        {
            using (var wb = Apply(input, output,
                new ExcelOperation("setColumnWidth", Sheet: "Data", Columns: "A:C", Width: 20),
                new ExcelOperation("setColumnWidth", Sheet: "Data", Columns: "D", AutoFit: true)))
            {
                var ws = wb.Worksheets[0];
                Assert.Equal(20d, ws.Columns["A"].WidthInCharacters, 0.5);
                Assert.Equal(20d, ws.Columns["C"].WidthInCharacters, 0.5);
                Assert.True(ws.Columns["D"].WidthInCharacters <= 60 && ws.Columns["D"].WidthInCharacters >= 50, $"D is {ws.Columns["D"].WidthInCharacters}");
            }

            var ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("setColumnWidth", Sheet: "Data", Columns: "A", Width: 20, AutoFit: true)));
            Assert.StartsWith("[invalid_operation] operations[0] (setColumnWidth): give exactly one", ex.Message);
            ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("setColumnWidth", Sheet: "Data", Columns: "A")));
            Assert.StartsWith("[invalid_operation] operations[0] (setColumnWidth): give exactly one", ex.Message);
            ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("setColumnWidth", Sheet: "Data", Columns: "A", Width: 300)));
            Assert.StartsWith("[invalid_operation] operations[0] (setColumnWidth): width 300", ex.Message);
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void FreezePanes_rows_rows_and_columns_and_unfreeze()
    {
        var input = Input();
        var rowsOnly = TempPath(".xlsx");
        var both = TempPath(".xlsx");
        var none = TempPath(".xlsx");
        try
        {
            Update(input, rowsOnly, new ExcelOperation("freezePanes", Sheet: "Data", Cell: "A2"));
            var xml = SheetXml(rowsOnly);
            Assert.Contains("state=\"frozen\"", xml);
            Assert.Contains("ySplit=\"1\"", xml);
            Assert.DoesNotContain("xSplit", xml);

            Update(input, both, new ExcelOperation("freezePanes", Sheet: "Data", Cell: "B3"));
            xml = SheetXml(both);
            Assert.Contains("ySplit=\"2\"", xml);
            Assert.Contains("xSplit=\"1\"", xml);

            Update(both, none, new ExcelOperation("freezePanes", Sheet: "Data", Cell: null));
            Assert.DoesNotContain("state=\"frozen\"", SheetXml(none));

            var ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("freezePanes", Sheet: "Data", Cell: "A1")));
            Assert.StartsWith("[invalid_operation] operations[0] (freezePanes): cell A1 freezes nothing", ex.Message);
        }
        finally { Delete(input, rowsOnly, both, none); }
    }

    [Fact]
    public void AutoFilter_apply_and_remove()
    {
        var input = Input();
        var filtered = TempPath(".xlsx");
        var cleared = TempPath(".xlsx");
        try
        {
            using (var wb = Apply(input, filtered, new ExcelOperation("autoFilter", Sheet: "Data", Range: "A1:C4")))
            {
                Assert.True(wb.Worksheets[0].AutoFilter.Enabled);
                Assert.Equal("A1:C4", wb.Worksheets[0].AutoFilter.Range.GetReferenceA1());
            }
            using (var wb = Apply(filtered, cleared, new ExcelOperation("autoFilter", Sheet: "Data", Range: null)))
                Assert.False(wb.Worksheets[0].AutoFilter.Enabled);
        }
        finally { Delete(input, filtered, cleared); }
    }
}
