using System.IO.Compression;
using McpOffice.Models;
using McpOffice.Services.Excel;
using McpOffice.Services.Excel.Write;
using ModelContextProtocol;
using static McpOffice.Tests.Excel.Write.WriteTestSupport;

namespace McpOffice.Tests.Excel.Write;

public class UpdateWorkbookMacroTests
{
    private static string CopyFixture(string name)
    {
        var copy = TempPath(".xlsm");
        File.Copy(TestFixtures.Path(name), copy);
        return copy;
    }

    private static readonly ExcelOperation AddSheet = new("addSheet", Name: "Added");

    [Fact]
    public void Xlsm_to_xlsm_keeps_vba_project()
    {
        var input = CopyFixture("sample-with-macros.xlsm");
        var output = TempPath(".xlsm");
        try
        {
            var result = Update(input, output, AddSheet);

            var modules = new ExcelWorkbookService().ExtractVba(output);
            Assert.True(modules.HasVbaProject);
            Assert.Equal(
                new ExcelWorkbookService().ExtractVba(input).Modules.Select(m => m.Name).Order(),
                modules.Modules.Select(m => m.Name).Order());
            Assert.Contains("Added", result.Sheets.Select(s => s.Name));
            Assert.DoesNotContain(result.Warnings, w => w.Contains("VBA"));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Xlsm_with_vba_to_xlsx_is_unsupported_format()
    {
        var input = CopyFixture("sample-with-macros.xlsm");
        var output = TempPath(".xlsx");
        try
        {
            var ex = Assert.Throws<McpException>(() => Update(input, output, AddSheet));
            Assert.StartsWith("[unsupported_format]", ex.Message);
            Assert.Contains("macros", ex.Message);
            Assert.Contains(".xlsm", ex.Message);
            Assert.False(File.Exists(output));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Xlsx_to_xlsm_is_allowed_and_a_clean_update_has_no_warnings()
    {
        var input = TestExcelWorkbooks.Create(wb => wb.Worksheets[0].Cells["A1"].Value = 1);
        var output = TempPath(".xlsm");
        try
        {
            var result = Update(input, output, AddSheet);
            Assert.Empty(result.Warnings);
            Assert.Equal(["Sheet1", "Added"], new ExcelWorkbookService().ListSheets(output).Select(s => s.Name));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void Update_reports_lost_macro_buttons_if_devexpress_drops_them()
    {
        var input = CopyFixture("synthetic-vba.xlsm");
        var output = TempPath(".xlsm");
        try
        {
            var result = Update(input, output, AddSheet);

            IReadOnlyDictionary<string, int> before, after;
            using (var zip = ZipFile.OpenRead(input)) before = PackagePartInventory.Take(zip);
            using (var zip = ZipFile.OpenRead(output)) after = PackagePartInventory.Take(zip);
            // Holds whether DevExpress keeps the parts or drops them: the warnings are exactly the inventory diff.
            Assert.Equal(PackagePartInventory.Compare(before, after), result.Warnings);
            Assert.True(after.ContainsKey(PackagePartInventory.VbaProject));
        }
        finally { Delete(input, output); }
    }
}
