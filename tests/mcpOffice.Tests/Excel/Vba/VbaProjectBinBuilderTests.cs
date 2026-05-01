using McpOffice.Services.Excel.Vba;
using OpenMcdf;

namespace McpOffice.Tests.Excel.Vba;

public class VbaProjectBinBuilderTests
{
    [Fact]
    public void Builder_output_round_trips_through_decompressor_and_dir_parser()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("Module1", "Module1", "Sub Hello()\r\nEnd Sub")
        ]);

        using var stream = new MemoryStream(bytes);
        using var root = RootStorage.Open(stream);
        Assert.True(root.TryOpenStorage("VBA", out var vba));
        Assert.NotNull(vba);

        Assert.True(vba!.TryOpenStream("dir", out var dirStream));
        Assert.NotNull(dirStream);

        byte[] dirBytes;
        using (dirStream)
        {
            dirBytes = new byte[dirStream!.Length];
            int read = 0;
            while (read < dirBytes.Length)
            {
                int n = dirStream.Read(dirBytes, read, dirBytes.Length - read);
                if (n == 0) break;
                read += n;
            }
        }

        var dirDecompressed = MsOvbaDecompressor.Decompress(dirBytes);
        var entries = VbaDirStreamParser.Parse(dirDecompressed);

        var entry = Assert.Single(entries);
        Assert.Equal("Module1", entry.Name);
        Assert.Equal("Module1", entry.StreamName);
        Assert.Equal((ushort)0x0021, entry.Type);
    }
}
