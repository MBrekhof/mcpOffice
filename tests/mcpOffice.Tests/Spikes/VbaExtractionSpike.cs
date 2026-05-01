using System.Text;
using OpenMcdf;

namespace McpOffice.Tests.Spikes;

// Throwaway spike for excel_extract_vba. Hardcoded local path; skipped on
// machines without the sample file. Output is dumped to disk so it survives
// xunit output buffering.
public class VbaExtractionSpike
{
    private const string VbaProjectPath = @"C:\temp\macro\vbaProject.bin";
    private const string OutputPath = @"C:\temp\macro\vba-spike-output.txt";

    [Fact]
    public void Probe_vbaProject_structure_and_dir_stream()
    {
        if (!File.Exists(VbaProjectPath))
        {
            return;
        }

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var writer = new StreamWriter(OutputPath, append: false);
        writer.WriteLine($"OpenMcdf probe of {VbaProjectPath}");
        writer.WriteLine($"File size: {new FileInfo(VbaProjectPath).Length} bytes");
        writer.WriteLine();

        using var root = RootStorage.OpenRead(VbaProjectPath);
        DumpStorage(root, "", writer);

        writer.WriteLine();
        writer.WriteLine("--- VBA / dir stream peek ---");
        TryPeekDir(root, writer);
    }

    private static void DumpStorage(Storage storage, string indent, StreamWriter writer)
    {
        foreach (var entry in storage.EnumerateEntries())
        {
            writer.WriteLine($"{indent}[{entry.Type}] {entry.Name}  ({entry.Length} bytes)");
            if (entry.Type == EntryType.Storage)
            {
                var child = storage.OpenStorage(entry.Name);
                DumpStorage(child, indent + "  ", writer);
            }
        }
    }

    private static void TryPeekDir(Storage root, StreamWriter writer)
    {
        if (!root.TryOpenStorage("VBA", out var vba) || vba is null)
        {
            writer.WriteLine("No 'VBA' storage at root.");
            return;
        }

        if (!vba.TryOpenStream("dir", out var dir) || dir is null)
        {
            writer.WriteLine("No 'dir' stream inside VBA storage.");
            return;
        }

        using (dir)
        {
            var compressed = new byte[dir.Length];
            int read = 0;
            while (read < compressed.Length)
            {
                int n = dir.Read(compressed, read, compressed.Length - read);
                if (n == 0) break;
                read += n;
            }

            writer.WriteLine($"dir stream compressed length: {compressed.Length} bytes");
            writer.WriteLine($"first 16 bytes: {BitConverter.ToString(compressed, 0, Math.Min(16, compressed.Length))}");

            try
            {
                var decompressed = MsOvbaDecompressor.Decompress(compressed);
                writer.WriteLine($"dir stream decompressed length: {decompressed.Length} bytes");
                writer.WriteLine($"first 64 bytes (hex): {BitConverter.ToString(decompressed, 0, Math.Min(64, decompressed.Length))}");

                writer.WriteLine();
                writer.WriteLine("Record walk (id size):");
                int rcCount = 0;
                int rj = 0;
                while (rj + 6 <= decompressed.Length && rcCount < 1500)
                {
                    ushort rid = BitConverter.ToUInt16(decompressed, rj);
                    uint rsize = BitConverter.ToUInt32(decompressed, rj + 2);
                    int actualPayload = (int)rsize;
                    if (rid == 0x0009) actualPayload = 6; // PROJECTVERSION: Reserved=4 but payload is 6
                    writer.WriteLine($"  pos={rj} id=0x{rid:X4} size={rsize} (payload={actualPayload})");
                    rj += 6 + actualPayload;
                    rcCount++;
                }
                writer.WriteLine($"(walked {rcCount} records, ended at pos {rj}/{decompressed.Length})");
                writer.WriteLine();

                var modules = ParseDirStream(decompressed);
                writer.WriteLine($"Modules discovered: {modules.Count}");
                foreach (var m in modules)
                {
                    writer.WriteLine($"  module: name='{m.Name}' streamName='{m.StreamName}' textOffset={m.TextOffset} type=0x{m.Type:X4}");
                }

                writer.WriteLine();
                writer.WriteLine("--- Source extraction sample ---");
                foreach (var sample in new[] { "Module2", "mdlWOM", "ThisWorkbook" })
                {
                    var info = modules.FirstOrDefault(x => x.Name == sample);
                    if (info is null) { writer.WriteLine($"[{sample}] not found"); continue; }
                    if (!vba.TryOpenStream(info.StreamName, out var modStream) || modStream is null)
                    { writer.WriteLine($"[{sample}] stream missing"); continue; }
                    using (modStream)
                    {
                        var modBytes = new byte[modStream.Length];
                        int rd = 0;
                        while (rd < modBytes.Length)
                        {
                            int n = modStream.Read(modBytes, rd, modBytes.Length - rd);
                            if (n == 0) break;
                            rd += n;
                        }
                        var compressedSource = modBytes.AsSpan((int)info.TextOffset).ToArray();
                        var sourceBytes = MsOvbaDecompressor.Decompress(compressedSource);
                        var source = Encoding.GetEncoding(1252).GetString(sourceBytes);
                        var preview = source.Length > 600 ? source[..600] + "..." : source;
                        writer.WriteLine();
                        writer.WriteLine($"[{sample}]  stream={modBytes.Length} bytes, textOffset={info.TextOffset}, decompressed source={sourceBytes.Length} bytes, char count={source.Length}");
                        writer.WriteLine("--- begin source preview ---");
                        writer.WriteLine(preview);
                        writer.WriteLine("--- end source preview ---");
                    }
                }
            }
            catch (Exception ex)
            {
                writer.WriteLine($"Decompression / parse failed: {ex.GetType().Name}: {ex.Message}");
            }
        }
    }

    private sealed record VbaModuleInfo(string Name, string StreamName, uint TextOffset, ushort Type);

    // MS-OVBA 2.3.4.2: dir stream is a sequence of (id u16, size u32, payload[size]) records.
    // Modules don't have a leading header; they are runs of records that start at MODULENAME
    // (0x0019) and end at Terminator (0x002B, size 0). Other 0x002B terminators (e.g. project
    // section terminators) appear too — we only emit when MODULENAME was seen in the run.
    private static List<VbaModuleInfo> ParseDirStream(byte[] data)
    {
        var result = new List<VbaModuleInfo>();
        var cp1252 = System.Text.Encoding.GetEncoding(1252);
        int i = 0;
        string? name = null;
        string? streamName = null;
        uint textOffset = 0;
        ushort type = 0;
        bool inModule = false;

        while (i + 6 <= data.Length)
        {
            ushort id = BitConverter.ToUInt16(data, i);
            uint size = BitConverter.ToUInt32(data, i + 2);
            int payloadStart = i + 6;
            int payloadLen = (int)size;
            if (id == 0x0009) payloadLen = 6; // PROJECTVERSION quirk: Reserved=4 but actual payload is 6
            if (payloadStart + payloadLen > data.Length) break;

            switch (id)
            {
                case 0x0019: // MODULENAME (MBCS) — also marks start of a module record run
                    inModule = true;
                    name = cp1252.GetString(data, payloadStart, payloadLen);
                    break;
                case 0x001A: // MODULESTREAMNAME (MBCS)
                    if (inModule) streamName = cp1252.GetString(data, payloadStart, payloadLen);
                    break;
                case 0x0031: // MODULEOFFSET (UInt32) — text offset within the module stream
                    if (inModule && payloadLen >= 4) textOffset = BitConverter.ToUInt32(data, payloadStart);
                    break;
                case 0x0021: // MODULETYPE procedural (size 0)
                case 0x0022: // MODULETYPE document/class (size 0)
                    if (inModule) type = id;
                    break;
                case 0x002B: // Terminator (size 0). May also terminate non-module sections.
                    if (inModule)
                    {
                        result.Add(new VbaModuleInfo(name ?? "", streamName ?? "", textOffset, type));
                        inModule = false;
                        name = null; streamName = null; textOffset = 0; type = 0;
                    }
                    break;
            }

            i = payloadStart + payloadLen;
        }

        return result;
    }
}

// MS-OVBA 2.4 RLE decompressor.
internal static class MsOvbaDecompressor
{
    public static byte[] Decompress(byte[] compressed)
    {
        if (compressed.Length < 1 || compressed[0] != 0x01)
            throw new InvalidDataException("Missing compressed-container signature byte 0x01");

        var output = new List<byte>(compressed.Length * 4);
        int pos = 1;
        while (pos < compressed.Length)
        {
            // Each chunk: 2-byte header, then up to 4096 bytes decompressed.
            if (pos + 2 > compressed.Length) break;
            ushort header = BitConverter.ToUInt16(compressed, pos);
            pos += 2;

            int chunkSize = (header & 0x0FFF) + 3; // includes header
            int chunkSig = (header >> 12) & 0x07;
            bool isCompressed = (header & 0x8000) != 0;
            if (chunkSig != 0b011)
                throw new InvalidDataException($"Bad chunk signature 0x{chunkSig:X} at pos {pos - 2}");

            int chunkEnd = pos + chunkSize - 2;
            if (chunkEnd > compressed.Length) chunkEnd = compressed.Length;

            int decompressedStartIndex = output.Count;

            if (!isCompressed)
            {
                // Raw chunk: next 4096 bytes are literal
                while (pos < chunkEnd) output.Add(compressed[pos++]);
            }
            else
            {
                while (pos < chunkEnd)
                {
                    byte flagByte = compressed[pos++];
                    for (int bit = 0; bit < 8 && pos < chunkEnd; bit++)
                    {
                        bool isCopy = (flagByte & (1 << bit)) != 0;
                        if (!isCopy)
                        {
                            output.Add(compressed[pos++]);
                        }
                        else
                        {
                            if (pos + 2 > chunkEnd) break;
                            ushort token = BitConverter.ToUInt16(compressed, pos);
                            pos += 2;

                            int diff = output.Count - decompressedStartIndex;
                            int lengthBits = CopyTokenLengthBits(diff);
                            int offsetBits = 16 - lengthBits;
                            int lengthMask = (1 << lengthBits) - 1;
                            int offsetMask = ~lengthMask & 0xFFFF;

                            int length = (token & lengthMask) + 3;
                            int offset = ((token & offsetMask) >> lengthBits) + 1;

                            int copyStart = output.Count - offset;
                            for (int j = 0; j < length; j++)
                                output.Add(output[copyStart + j]);
                        }
                    }
                }
            }
        }
        return output.ToArray();
    }

    // MS-OVBA 2.4.1.3.19.1: number of length bits = ceiling(log2(decompressed-so-far)),
    // clamped to [4, 12].
    private static int CopyTokenLengthBits(int decompressedLengthSoFar)
    {
        if (decompressedLengthSoFar <= 16) return 12;
        if (decompressedLengthSoFar <= 32) return 11;
        if (decompressedLengthSoFar <= 64) return 10;
        if (decompressedLengthSoFar <= 128) return 9;
        if (decompressedLengthSoFar <= 256) return 8;
        if (decompressedLengthSoFar <= 512) return 7;
        if (decompressedLengthSoFar <= 1024) return 6;
        if (decompressedLengthSoFar <= 2048) return 5;
        return 4;
    }
}
