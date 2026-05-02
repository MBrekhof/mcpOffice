using McpOffice.Services.Excel;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel;

public class GetStructureTests
{
    [Fact]
    public void GetStructure_returns_sheets_with_formula_counts_and_defined_names()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            var data = workbook.Worksheets[0];
            data.Name = "Data";
            data.Cells["A1"].Value = 1;
            data.Cells["A2"].Value = 2;
            data.Cells["B1"].Formula = "=A1*2";
            data.Cells["B2"].Formula = "=A2*2";

            var notes = workbook.Worksheets.Add("Notes");
            notes.Cells["A1"].Value = "hello";

            workbook.DefinedNames.Add("TaxRate", "=0.21");
        });

        try
        {
            var structure = new ExcelWorkbookService().GetStructure(
                path,
                includeSheets: true,
                includeFormulaCounts: true,
                includeDefinedNames: true);

            Assert.Equal(2, structure.SheetCount);
            Assert.Equal(1, structure.DefinedNameCount);

            Assert.NotNull(structure.Sheets);
            Assert.Equal(2, structure.Sheets!.Count);

            var data = structure.Sheets!.Single(s => s.Name == "Data");
            Assert.Equal(0, data.Index);
            Assert.True(data.Visible);
            Assert.Equal("A1:B2", data.UsedRange);
            Assert.Equal(2, data.FormulaCount);

            var notes = structure.Sheets!.Single(s => s.Name == "Notes");
            Assert.Equal(0, notes.FormulaCount);

            Assert.NotNull(structure.DefinedNames);
            var taxRate = Assert.Single(structure.DefinedNames!);
            Assert.Equal("TaxRate", taxRate.Name);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void GetStructure_omits_collections_when_disabled()
    {
        var path = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Cells["A1"].Value = 1;
            workbook.DefinedNames.Add("TaxRate", "=0.21");
        });

        try
        {
            var structure = new ExcelWorkbookService().GetStructure(
                path,
                includeSheets: false,
                includeFormulaCounts: false,
                includeDefinedNames: false);

            Assert.Equal(1, structure.SheetCount);
            Assert.Equal(1, structure.DefinedNameCount);
            Assert.Null(structure.Sheets);
            Assert.Null(structure.DefinedNames);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void GetStructure_throws_file_not_found_for_missing_workbook()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"mcpoffice-missing-{Guid.NewGuid():N}.xlsx");

        var ex = Assert.Throws<McpException>(() => new ExcelWorkbookService().GetStructure(
            missing,
            includeSheets: true,
            includeFormulaCounts: true,
            includeDefinedNames: true));

        Assert.Contains("file_not_found", ex.Message);
    }
}
