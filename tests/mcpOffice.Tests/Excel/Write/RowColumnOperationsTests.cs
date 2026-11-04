using McpOffice.Models;
using ModelContextProtocol;
using static McpOffice.Tests.Excel.Write.WriteTestSupport;

namespace McpOffice.Tests.Excel.Write;

public class RowColumnOperationsTests
{
    /// <summary>A1..A3 = r1,r2,r3; B2=10, B3=20, B4 "=SUM(B2:B3)"; C1=c1.</summary>
    private static string Input() => TestExcelWorkbooks.Create(wb =>
    {
        var ws = wb.Worksheets[0];
        ws.Name = "Data";
        ws.Cells["A1"].Value = "r1";
        ws.Cells["A2"].Value = "r2";
        ws.Cells["A3"].Value = "r3";
        ws.Cells["B2"].Value = 10;
        ws.Cells["B3"].Value = 20;
        ws.Cells["B4"].Formula = "=SUM(B2:B3)";
        ws.Cells["C1"].Value = "c1";
        wb.Calculate();
    });

    [Fact]
    public void InsertRows_shifts_down_and_adjusts_formulas()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            Update(input, output, new ExcelOperation("insertRows", Sheet: "Data", Row: 2, Count: 2));
            var data = Read(output, "Data", "A1:B6");
            Assert.Equal("r1", data.Cell("A1").Value);
            Assert.Equal("empty", data.Cell("A2").ValueType);
            Assert.Equal("empty", data.Cell("A3").ValueType);
            Assert.Equal("r2", data.Cell("A4").Value);
            Assert.Equal("=SUM(B4:B5)", data.Cell("B6").Formula);
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void DeleteRows_shifts_up()
    {
        var input = Input();
        var output = TempPath(".xlsx");
        try
        {
            Update(input, output, new ExcelOperation("deleteRows", Sheet: "Data", Row: 1));
            var data = Read(output, "Data", "A1:B3");
            Assert.Equal("r2", data.Cell("A1").Value);
            Assert.Equal("r3", data.Cell("A2").Value);
            Assert.Equal("=SUM(B1:B2)", data.Cell("B3").Formula);
        }
        finally { Delete(input, output); }
    }

    [Theory]
    [InlineData("B")]
    [InlineData("2")]
    public void InsertColumns_and_deleteColumns(string column)
    {
        var input = Input();
        var inserted = TempPath(".xlsx");
        var deleted = TempPath(".xlsx");
        try
        {
            Update(input, inserted, new ExcelOperation("insertColumns", Sheet: "Data", Column: column, Count: 2));
            var data = Read(inserted, "Data", "A1:E4");
            Assert.Equal("empty", data.Cell("B2").ValueType);
            Assert.Equal(10d, data.Cell("D2").Value);
            Assert.Equal("c1", data.Cell("E1").Value);
            Assert.Equal("=SUM(D2:D3)", data.Cell("D4").Formula);

            Update(inserted, deleted, new ExcelOperation("deleteColumns", Sheet: "Data", Column: column, Count: 2));
            data = Read(deleted, "Data", "A1:C4");
            Assert.Equal(10d, data.Cell("B2").Value);
            Assert.Equal("c1", data.Cell("C1").Value);
        }
        finally { Delete(input, inserted, deleted); }
    }

    [Fact]
    public void Bad_row_and_count_are_invalid_operation()
    {
        var input = Input();
        try
        {
            var ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("insertRows", Sheet: "Data", Row: 2, Count: 0)));
            Assert.StartsWith("[invalid_operation] operations[0] (insertRows): count 0", ex.Message);
            ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("deleteRows", Sheet: "Data", Row: 2, Count: -1)));
            Assert.StartsWith("[invalid_operation] operations[0] (deleteRows): count -1", ex.Message);
            ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("insertRows", Sheet: "Data")));
            Assert.StartsWith("[invalid_operation] operations[0] (insertRows): row is required", ex.Message);
            ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("insertRows", Sheet: "Data", Row: 0)));
            Assert.StartsWith("[invalid_operation] operations[0] (insertRows): row 0", ex.Message);
            ex = Assert.Throws<McpException>(() => Update(input, TempPath(".xlsx"), new ExcelOperation("insertColumns", Sheet: "Data")));
            Assert.StartsWith("[invalid_operation] operations[0] (insertColumns): column is required", ex.Message);
        }
        finally { Delete(input); }
    }
}
