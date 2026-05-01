using System.Runtime.CompilerServices;
using System.Text;
using OpenMcdf;

namespace McpOffice.Tests.Excel.Vba;

internal sealed record ModuleSpec(
    string Name,
    string StreamName,
    string Source,
    bool IsDocumentModule = false,
    string? UnicodeName = null);

// Constructs a synthetic vbaProject.bin (OLE compound file) for unit-testing
// VbaProjectReader.ReadVbaProjectBin. Emits the minimum structure the production
// reader needs:
//   VBA/dir         - MS-OVBA-compressed dir stream with project records + per-module
//                     records (MODULENAME / MODULESTREAMNAME / MODULEOFFSET / MODULETYPE
//                     / Terminator), plus a top-level Terminator.
//   VBA/<stream>    - MS-OVBA-compressed source for each module.
//
// Compression always uses literal-only compressed-mode chunks (flag bytes = 0x00).
// The production decompressor handles literal-only and copy-token chunks identically.
internal static class VbaProjectBinBuilder
{
    [ModuleInitializer]
    internal static void RegisterCp1252() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public static byte[] Build(IReadOnlyList<ModuleSpec> modules)
    {
        var dir = BuildDirStream(modules);
        var dirCompressed = CompressLiteralsOnly(dir);

        var moduleStreams = modules.ToDictionary(
            m => m.StreamName,
            m => CompressLiteralsOnly(Encoding.GetEncoding(1252).GetBytes(m.Source)));

        return BuildOleCompoundFile(dirCompressed, moduleStreams);
    }

    private static byte[] BuildDirStream(IReadOnlyList<ModuleSpec> modules)
    {
        using var ms = new MemoryStream();
        var cp1252 = Encoding.GetEncoding(1252);

        // PROJECTSYSKIND (0x0001) — single UInt32 payload.
        WriteRecord(ms, 0x0001, sizeField: 4, payload: BitConverter.GetBytes(1u));

        // PROJECTVERSION (0x0009) — sizeField says 4, payload is 6. Production parser
        // must special-case this; emitting it here exercises that path.
        WriteRecord(ms, 0x0009, sizeField: 4, payload: [0x01, 0x00, 0x00, 0x00, 0x02, 0x00]);

        foreach (var m in modules)
        {
            var nameBytes = cp1252.GetBytes(m.Name);
            WriteRecord(ms, 0x0019, (uint)nameBytes.Length, nameBytes);

            if (m.UnicodeName is not null)
            {
                var unicodeNameBytes = Encoding.Unicode.GetBytes(m.UnicodeName);
                WriteRecord(ms, 0x0047, (uint)unicodeNameBytes.Length, unicodeNameBytes);
            }

            var streamBytes = cp1252.GetBytes(m.StreamName);
            WriteRecord(ms, 0x001A, (uint)streamBytes.Length, streamBytes);

            var unicodeStreamBytes = Encoding.Unicode.GetBytes(m.StreamName);
            WriteRecord(ms, 0x0032, (uint)unicodeStreamBytes.Length, unicodeStreamBytes);

            WriteRecord(ms, 0x0031, sizeField: 4, payload: BitConverter.GetBytes(0u));

            ushort moduleType = m.IsDocumentModule ? (ushort)0x0022 : (ushort)0x0021;
            WriteRecord(ms, moduleType, sizeField: 0, payload: []);

            WriteRecord(ms, 0x002B, sizeField: 0, payload: []);
        }

        WriteRecord(ms, 0x002B, sizeField: 0, payload: []);
        return ms.ToArray();
    }

    private static void WriteRecord(Stream s, ushort id, uint sizeField, byte[] payload)
    {
        s.Write(BitConverter.GetBytes(id), 0, 2);
        s.Write(BitConverter.GetBytes(sizeField), 0, 4);
        if (payload.Length > 0) s.Write(payload, 0, payload.Length);
    }

    // MS-OVBA 2.4.1 — literal-only compressed-mode encoding.
    // Layout: 0x01 signature, then chunks of up to 4096 source bytes each:
    //   chunk header (2 bytes LE): (chunkSize-3) | (0b011 << 12) | (1 << 15)
    //   for each group of up to 8 source bytes: 1 flag byte = 0x00, then the literals.
    private static byte[] CompressLiteralsOnly(byte[] data)
    {
        using var ms = new MemoryStream();
        ms.WriteByte(0x01);

        if (data.Length == 0) return ms.ToArray();

        int offset = 0;
        while (offset < data.Length)
        {
            int segmentLen = Math.Min(4096, data.Length - offset);
            int groupCount = (segmentLen + 7) / 8;
            int payloadLen = groupCount + segmentLen;
            int chunkSize = 2 + payloadLen;
            ushort header = (ushort)(((chunkSize - 3) & 0x0FFF) | (0b011 << 12) | (1 << 15));
            ms.WriteByte((byte)(header & 0xFF));
            ms.WriteByte((byte)((header >> 8) & 0xFF));

            int written = 0;
            while (written < segmentLen)
            {
                ms.WriteByte(0x00);
                int groupLen = Math.Min(8, segmentLen - written);
                ms.Write(data, offset + written, groupLen);
                written += groupLen;
            }

            offset += segmentLen;
        }

        return ms.ToArray();
    }

    private static byte[] BuildOleCompoundFile(byte[] dirCompressed, IReadOnlyDictionary<string, byte[]> moduleStreams)
    {
        var ms = new MemoryStream();
        using (var root = RootStorage.Create(ms, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
        {
            var vba = root.CreateStorage("VBA");

            using (var dirStream = vba.CreateStream("dir"))
            {
                dirStream.Write(dirCompressed, 0, dirCompressed.Length);
            }

            foreach (var (streamName, payload) in moduleStreams)
            {
                using var ms2 = vba.CreateStream(streamName);
                ms2.Write(payload, 0, payload.Length);
            }
        }

        return ms.ToArray();
    }
}
