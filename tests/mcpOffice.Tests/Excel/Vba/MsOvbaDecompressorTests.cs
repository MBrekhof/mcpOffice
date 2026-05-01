using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

// Hand-crafted MS-OVBA inputs per spec section 2.4.1. Chunk header format:
//   bits 0-11 = chunkSize - 3 (chunkSize includes the 2-byte header)
//   bits 12-14 = signature 0b011
//   bit 15     = compressed flag (1 = compressed-mode chunk)
public class MsOvbaDecompressorTests
{
    [Fact]
    public void Throws_when_signature_byte_missing()
    {
        var bad = new byte[] { 0x00, 0x00, 0x00 };
        Assert.Throws<InvalidDataException>(() => MsOvbaDecompressor.Decompress(bad));
    }

    [Fact]
    public void Throws_when_input_is_empty()
    {
        Assert.Throws<InvalidDataException>(() => MsOvbaDecompressor.Decompress([]));
    }

    [Fact]
    public void Decompresses_single_chunk_of_eight_literals()
    {
        // 8 literal bytes "ABCDEFGH". chunkSize = 2 header + 1 flag + 8 literals = 11.
        // header raw = (11-3) | 0x3000 | 0x8000 = 0xB008 → LE bytes 0x08 0xB0.
        var bytes = new byte[]
        {
            0x01,
            0x08, 0xB0,
            0x00,
            0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48
        };

        var result = MsOvbaDecompressor.Decompress(bytes);

        Assert.Equal("ABCDEFGH", System.Text.Encoding.ASCII.GetString(result));
    }

    [Fact]
    public void Decompresses_chunk_spanning_multiple_flag_bytes()
    {
        // 10 literals "abcdefghij" → 2 flag-byte groups (8 + 2).
        // chunkSize = 2 header + 1 flag + 8 literals + 1 flag + 2 literals = 14.
        // header raw = (14-3) | 0x3000 | 0x8000 = 0xB00B → LE 0x0B 0xB0.
        var bytes = new byte[]
        {
            0x01,
            0x0B, 0xB0,
            0x00, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68,
            0x00, 0x69, 0x6A
        };

        var result = MsOvbaDecompressor.Decompress(bytes);

        Assert.Equal("abcdefghij", System.Text.Encoding.ASCII.GetString(result));
    }

    [Fact]
    public void Decompresses_copy_token_for_repeated_run()
    {
        // Literals "AB" then a copy-token: length=3 offset=1 → copies 'B' 3× → "ABBBB".
        // At decompressed-so-far=2 (≤16) lengthBits=12. length=3 encoded as 0; offset=1 encoded as 0.
        // Token = 0x0000.
        // Flag byte: bit0=0 literal, bit1=0 literal, bit2=1 copy → 0b00000100 = 0x04.
        // chunkSize = 2 header + 1 flag + 1 ('A') + 1 ('B') + 2 (token) = 7.
        // header raw = (7-3) | 0x3000 | 0x8000 = 0xB004 → LE 0x04 0xB0.
        var bytes = new byte[]
        {
            0x01,
            0x04, 0xB0,
            0x04,
            0x41, 0x42,
            0x00, 0x00
        };

        var result = MsOvbaDecompressor.Decompress(bytes);

        Assert.Equal("ABBBB", System.Text.Encoding.ASCII.GetString(result));
    }

    [Fact]
    public void Throws_on_bad_chunk_signature()
    {
        // chunk header signature nibble forced to 0 (spec requires 0b011).
        // Raw header = 0x0008 (bits 12-14 = 0). LE: 0x08 0x00.
        var bytes = new byte[]
        {
            0x01,
            0x08, 0x00,
            0x00,
            0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48
        };

        Assert.Throws<InvalidDataException>(() => MsOvbaDecompressor.Decompress(bytes));
    }

    [Fact]
    public void Decompresses_uncompressed_chunk()
    {
        // Uncompressed-mode chunk (bit 15 clear). Per MS-OVBA, an uncompressed chunk
        // carries up to 4096 raw bytes. Use 5 here to keep the test small; the
        // decompressor reads to chunkEnd, not to a fixed 4096 boundary.
        // chunkSize = 2 header + 5 raw = 7. header raw = (7-3) | 0x3000 | 0x0000 = 0x3004
        // LE: 0x04 0x30.
        var bytes = new byte[]
        {
            0x01,
            0x04, 0x30,
            0x48, 0x65, 0x6C, 0x6C, 0x6F   // "Hello"
        };

        var result = MsOvbaDecompressor.Decompress(bytes);

        Assert.Equal("Hello", System.Text.Encoding.ASCII.GetString(result));
    }
}
