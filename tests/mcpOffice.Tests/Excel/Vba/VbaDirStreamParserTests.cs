using System.Runtime.CompilerServices;
using System.Text;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaDirStreamParserTests
{
    [ModuleInitializer]
    internal static void RegisterCp1252() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public void Walks_past_PROJECTVERSION_record_with_real_payload_size_6()
    {
        // PROJECTVERSION (id=0x0009) declares size=4 but the spec puts a 6-byte payload
        // there. The parser must add 6, not 4, or it desynchronises and returns no
        // modules from the stream that follows.
        var stream = BuildDirStream(
            Record(0x0009, sizeField: 4, payload: [0x01, 0x00, 0x00, 0x00, 0x02, 0x00]),
            Record(0x0019, sizeField: 7, payload: AsMbcs("Module1")),
            Record(0x001A, sizeField: 7, payload: AsMbcs("Module1")),
            Record(0x0031, sizeField: 4, payload: BitConverter.GetBytes(0u)),
            Record(0x0021, sizeField: 0, payload: []),
            Record(0x002B, sizeField: 0, payload: []));

        var modules = VbaDirStreamParser.Parse(stream);

        var m = Assert.Single(modules);
        Assert.Equal("Module1", m.Name);
        Assert.Equal((ushort)0x0021, m.Type);
    }

    [Fact]
    public void Prefers_unicode_module_name_when_both_are_present()
    {
        var stream = BuildDirStream(
            Record(0x0019, sizeField: 6, payload: AsMbcs("ModABC")),
            Record(0x0047, sizeField: 12, payload: AsUtf16Le("ΜοδΑΒΓ")),
            Record(0x001A, sizeField: 6, payload: AsMbcs("ModABC")),
            Record(0x0032, sizeField: 12, payload: AsUtf16Le("ΜοδΑΒΓ")),
            Record(0x0031, sizeField: 4, payload: BitConverter.GetBytes(0u)),
            Record(0x0022, sizeField: 0, payload: []),
            Record(0x002B, sizeField: 0, payload: []));

        var modules = VbaDirStreamParser.Parse(stream);

        var m = Assert.Single(modules);
        Assert.Equal("ΜοδΑΒΓ", m.Name);
        Assert.Equal("ΜοδΑΒΓ", m.StreamName);
        Assert.Equal((ushort)0x0022, m.Type);
    }

    [Fact]
    public void Returns_modules_in_order()
    {
        var stream = BuildDirStream(
            Record(0x0019, sizeField: 4, payload: AsMbcs("ModA")),
            Record(0x001A, sizeField: 4, payload: AsMbcs("ModA")),
            Record(0x0021, sizeField: 0, payload: []),
            Record(0x002B, sizeField: 0, payload: []),
            Record(0x0019, sizeField: 4, payload: AsMbcs("ModB")),
            Record(0x001A, sizeField: 4, payload: AsMbcs("ModB")),
            Record(0x0021, sizeField: 0, payload: []),
            Record(0x002B, sizeField: 0, payload: []));

        var modules = VbaDirStreamParser.Parse(stream);

        Assert.Equal(["ModA", "ModB"], modules.Select(m => m.Name).ToArray());
    }

    [Fact]
    public void Bare_terminator_does_not_emit_phantom_module()
    {
        // A 0x002B terminator before any MODULENAME (e.g. project-section terminator)
        // must not produce an empty module entry.
        var stream = BuildDirStream(
            Record(0x0001, sizeField: 4, payload: BitConverter.GetBytes(0u)),  // PROJECTSYSKIND
            Record(0x002B, sizeField: 0, payload: []));

        var modules = VbaDirStreamParser.Parse(stream);

        Assert.Empty(modules);
    }

    [Fact]
    public void Captures_textOffset_from_MODULEOFFSET_record()
    {
        const uint expectedOffset = 0x1234;
        var stream = BuildDirStream(
            Record(0x0019, sizeField: 4, payload: AsMbcs("ModA")),
            Record(0x001A, sizeField: 4, payload: AsMbcs("ModA")),
            Record(0x0031, sizeField: 4, payload: BitConverter.GetBytes(expectedOffset)),
            Record(0x0021, sizeField: 0, payload: []),
            Record(0x002B, sizeField: 0, payload: []));

        var modules = VbaDirStreamParser.Parse(stream);

        Assert.Equal(expectedOffset, modules.Single().TextOffset);
    }

    private static byte[] BuildDirStream(params byte[][] records)
    {
        using var ms = new MemoryStream();
        foreach (var rec in records) ms.Write(rec, 0, rec.Length);
        return ms.ToArray();
    }

    private static byte[] Record(ushort id, uint sizeField, byte[] payload)
    {
        var buffer = new byte[6 + payload.Length];
        BitConverter.GetBytes(id).CopyTo(buffer, 0);
        BitConverter.GetBytes(sizeField).CopyTo(buffer, 2);
        payload.CopyTo(buffer, 6);
        return buffer;
    }

    private static byte[] AsMbcs(string text) => Encoding.GetEncoding(1252).GetBytes(text);
    private static byte[] AsUtf16Le(string text) => Encoding.Unicode.GetBytes(text);
}
