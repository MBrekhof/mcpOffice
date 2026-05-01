using System.Text;

namespace McpOffice.Services.Excel.Vba;

// MS-OVBA 2.3.4.2 — dir stream is a sequence of (id u16 | size u32 | payload[size])
// records. A module record run is delimited by MODULENAME (0x0019) or MODULENAMEUNICODE
// (0x0047) at the start and a Terminator (0x002B, size 0) at the end. Other 0x002B
// terminators (project-section ends) appear too, so we only emit when at least one
// MODULENAME-flavored record was seen in the run.
internal static class VbaDirStreamParser
{
    private const ushort IdProjectVersion = 0x0009;
    private const ushort IdModuleName = 0x0019;
    private const ushort IdModuleNameUnicode = 0x0047;
    private const ushort IdModuleStreamName = 0x001A;
    private const ushort IdModuleStreamNameUnicode = 0x0032;
    private const ushort IdModuleOffset = 0x0031;
    private const ushort IdModuleTypeProcedural = 0x0021;
    private const ushort IdModuleTypeDocument = 0x0022;
    private const ushort IdTerminator = 0x002B;

    public static IReadOnlyList<VbaModuleEntry> Parse(byte[] decompressedDirStream)
    {
        var result = new List<VbaModuleEntry>();
        var cp1252 = Encoding.GetEncoding(1252);

        int i = 0;
        string? mbcsName = null;
        string? unicodeName = null;
        string? mbcsStreamName = null;
        string? unicodeStreamName = null;
        uint textOffset = 0;
        ushort type = 0;
        bool inModule = false;

        while (i + 6 <= decompressedDirStream.Length)
        {
            ushort id = BitConverter.ToUInt16(decompressedDirStream, i);
            uint size = BitConverter.ToUInt32(decompressedDirStream, i + 2);
            int payloadStart = i + 6;
            int payloadLen = (int)size;

            // PROJECTVERSION quirk: the size field reads as 4 but the actual payload
            // is 6 bytes (UInt32 major + UInt16 minor). Walking past it without this
            // special-case throws the parser off by 2 bytes for the rest of the stream.
            if (id == IdProjectVersion) payloadLen = 6;

            if (payloadStart + payloadLen > decompressedDirStream.Length) break;

            switch (id)
            {
                case IdModuleName:
                    inModule = true;
                    mbcsName = cp1252.GetString(decompressedDirStream, payloadStart, payloadLen);
                    break;
                case IdModuleNameUnicode:
                    inModule = true;
                    unicodeName = Encoding.Unicode.GetString(decompressedDirStream, payloadStart, payloadLen);
                    break;
                case IdModuleStreamName:
                    if (inModule)
                        mbcsStreamName = cp1252.GetString(decompressedDirStream, payloadStart, payloadLen);
                    break;
                case IdModuleStreamNameUnicode:
                    if (inModule)
                        unicodeStreamName = Encoding.Unicode.GetString(decompressedDirStream, payloadStart, payloadLen);
                    break;
                case IdModuleOffset:
                    if (inModule && payloadLen >= 4)
                        textOffset = BitConverter.ToUInt32(decompressedDirStream, payloadStart);
                    break;
                case IdModuleTypeProcedural:
                case IdModuleTypeDocument:
                    if (inModule) type = id;
                    break;
                case IdTerminator:
                    if (inModule)
                    {
                        var name = unicodeName ?? mbcsName ?? "";
                        var streamName = unicodeStreamName ?? mbcsStreamName ?? "";
                        result.Add(new VbaModuleEntry(name, streamName, textOffset, type));

                        inModule = false;
                        mbcsName = null;
                        unicodeName = null;
                        mbcsStreamName = null;
                        unicodeStreamName = null;
                        textOffset = 0;
                        type = 0;
                    }
                    break;
            }

            i = payloadStart + payloadLen;
        }

        return result;
    }
}
