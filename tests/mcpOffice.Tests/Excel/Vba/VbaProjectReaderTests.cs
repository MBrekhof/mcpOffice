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

    [Fact]
    public void Reads_modules_from_real_excel_fixture()
    {
        // End-to-end: zip → vbaProject.bin → real MS-OVBA compressed chunks → modules.
        // The synthetic builder uses literal-only compressed chunks; this test ensures
        // we still parse Excel's actual (copy-token) compressed output.
        var path = TestFixtures.Path("sample-with-macros.xlsm");
        var project = new VbaProjectReader().Read(path);

        Assert.True(project.HasVbaProject);
        Assert.Contains(project.Modules, m => m.Name == "Module1" && m.Kind == "standardModule");
        Assert.Contains(project.Modules, m => m.Name == "ThisWorkbook" && m.Kind == "documentModule");

        var module1 = project.Modules.Single(m => m.Name == "Module1");
        Assert.Contains("Sub Hello", module1.Code);
        Assert.True(module1.LineCount > 0);
    }

    [Fact]
    public void Classifies_userform_module_when_form_storage_exists()
    {
        var formNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "UserForm1" };
        var bytes = VbaProjectBinBuilder.Build(
            [
                new ModuleSpec("Module1", "Module1", "Sub Foo()\nEnd Sub"),
                // UserForms use the class-type module record (0x0022, IsDocumentModule=true in the
                // builder) in the dir stream — the form storage at root is what distinguishes them
                // from plain classModules or documentModules.
                new ModuleSpec("UserForm1", "UserForm1", "Private Sub UserForm_Initialize()\nEnd Sub",
                    IsDocumentModule: true)
            ],
            formModuleNames: formNames);

        using var ms = new MemoryStream(bytes);
        var project = new VbaProjectReader().ReadVbaProjectBin(ms, "synthetic");

        Assert.True(project.HasVbaProject);
        Assert.Contains(project.Modules, m => m.Name == "Module1" && m.Kind == "standardModule");
        Assert.Contains(project.Modules, m => m.Name == "UserForm1" && m.Kind == "userForm");
    }

    [Fact(Skip = "needs locked-project sample — see SESSION_HANDOFF.md Open Question #1")]
    public void Throws_vba_project_locked_for_protected_project() { }
}
