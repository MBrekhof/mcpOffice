using McpOffice.Services.Excel.Vba;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class VbaProjectReaderTests
{
    [Fact]
    public void Reads_single_standard_module()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("Module1", "Module1", "Sub Hello()\r\n  Debug.Print \"hi\"\r\nEnd Sub")
        ]);

        var project = new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(bytes), "<synthetic>");

        Assert.True(project.HasVbaProject);
        var m = Assert.Single(project.Modules);
        Assert.Equal("Module1", m.Name);
        Assert.Equal("standardModule", m.Kind);
        Assert.Contains("Sub Hello", m.Code);
        Assert.Equal(3, m.LineCount);
    }

    [Fact]
    public void Reads_document_module_with_correct_kind()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("ThisWorkbook", "ThisWorkbook",
                "Private Sub Workbook_Open()\r\nEnd Sub",
                IsDocumentModule: true)
        ]);

        var project = new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(bytes), "<synthetic>");
        Assert.Equal("documentModule", project.Modules.Single().Kind);
    }

    [Fact]
    public void Returns_modules_in_dir_stream_order()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("ModA", "ModA", "' a"),
            new ModuleSpec("ModB", "ModB", "' b"),
            new ModuleSpec("ThisWorkbook", "ThisWorkbook", "' wb", IsDocumentModule: true)
        ]);

        var project = new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(bytes), "<synthetic>");

        Assert.Equal(["ModA", "ModB", "ThisWorkbook"], project.Modules.Select(m => m.Name).ToArray());
    }

    [Fact]
    public void Prefers_unicode_module_name_when_present()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("ModABC", "ModABC", "' x", UnicodeName: "ΜοδΑΒΓ")
        ]);

        var project = new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(bytes), "<synthetic>");

        Assert.Equal("ΜοδΑΒΓ", project.Modules.Single().Name);
    }

    [Fact]
    public void Throws_vba_parse_error_for_corrupt_input()
    {
        var corrupt = new byte[] { 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07 };

        var ex = Assert.Throws<McpException>(() =>
            new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(corrupt), "<synthetic>"));

        Assert.Contains("vba_parse_error", ex.Message);
        Assert.Contains("<synthetic>", ex.Message);
    }

    [Fact(Skip = "needs locked-project sample — see SESSION_HANDOFF.md Open Question #1")]
    public void Throws_vba_project_locked_for_protected_project() { }
}
