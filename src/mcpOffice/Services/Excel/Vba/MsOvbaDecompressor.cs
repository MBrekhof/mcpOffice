namespace McpOffice.Services.Excel.Vba;

// MS-OVBA 2.4 RLE decompressor for VBA compressed-container streams.
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
