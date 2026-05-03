using System.IO.Compression;
using System.Text;
using McpOffice.Models;
using ModelContextProtocol;
using OpenMcdf;

namespace McpOffice.Services.Excel.Vba;

internal sealed class VbaProjectReader
{
    private const string VbaProjectEntryName = "xl/vbaProject.bin";
    private const string VbaStorageName = "VBA";
    private const string DirStreamName = "dir";

    static VbaProjectReader()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public ExcelVbaProject Read(string xlsmPath)
    {
        try
        {
            using var zip = ZipFile.OpenRead(xlsmPath);
            var entry = zip.GetEntry(VbaProjectEntryName);
            if (entry is null)
            {
                return new ExcelVbaProject(HasVbaProject: false, Modules: []);
            }

            using var ms = new MemoryStream();
            using (var s = entry.Open()) s.CopyTo(ms);
            ms.Position = 0;
            return ReadVbaProjectBin(ms, xlsmPath);
        }
        catch (McpException) { throw; }
        catch (Exception ex)
        {
            throw ToolError.VbaParseError(xlsmPath, ex.Message);
        }
    }

    public ExcelVbaProject ReadVbaProjectBin(Stream vbaProjectBin, string sourceLabel)
    {
        try
        {
            using var root = RootStorage.Open(vbaProjectBin);
            if (!root.TryOpenStorage(VbaStorageName, out var vba) || vba is null)
            {
                throw ToolError.VbaProjectLocked(sourceLabel);
            }

            var dirBytes = ReadDirStream(vba, sourceLabel);
            byte[] dirDecompressed;
            try
            {
                dirDecompressed = MsOvbaDecompressor.Decompress(dirBytes);
            }
            catch (InvalidDataException ex)
            {
                throw ToolError.VbaParseError(sourceLabel, $"dir decompression failed: {ex.Message}");
            }

            var entries = VbaDirStreamParser.Parse(dirDecompressed);
            if (entries.Count == 0)
            {
                throw ToolError.VbaProjectLocked(sourceLabel);
            }

            // Build a case-insensitive set of storage names that indicate UserForm modules.
            // UserForms have a same-named sub-storage (holding layout streams like f/o).
            // Excel places these at the root level; some versions put them inside VBA.
            // We check both locations to be safe.
            var formStorageNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var info in root.EnumerateEntries())
            {
                if (info.Type == EntryType.Storage &&
                    !string.Equals(info.Name, VbaStorageName, StringComparison.OrdinalIgnoreCase))
                {
                    formStorageNames.Add(info.Name);
                }
            }
            foreach (var info in vba.EnumerateEntries())
            {
                if (info.Type == EntryType.Storage)
                    formStorageNames.Add(info.Name);
            }

            var modules = new List<ExcelVbaModule>(entries.Count);
            var cp1252 = Encoding.GetEncoding(1252);

            foreach (var entry in entries)
            {
                if (string.IsNullOrEmpty(entry.StreamName) ||
                    !vba.TryOpenStream(entry.StreamName, out var modStream) ||
                    modStream is null)
                {
                    throw ToolError.VbaParseError(sourceLabel,
                        $"module stream '{entry.StreamName}' missing for module '{entry.Name}'");
                }

                byte[] modBytes;
                using (modStream)
                {
                    modBytes = new byte[modStream.Length];
                    int read = 0;
                    while (read < modBytes.Length)
                    {
                        int n = modStream.Read(modBytes, read, modBytes.Length - read);
                        if (n == 0) break;
                        read += n;
                    }
                }

                if (entry.TextOffset > modBytes.Length)
                {
                    throw ToolError.VbaParseError(sourceLabel,
                        $"textOffset {entry.TextOffset} exceeds module stream length {modBytes.Length} for '{entry.Name}'");
                }

                var compressedSource = modBytes.AsSpan((int)entry.TextOffset).ToArray();
                byte[] sourceBytes;
                try
                {
                    sourceBytes = MsOvbaDecompressor.Decompress(compressedSource);
                }
                catch (InvalidDataException ex)
                {
                    throw ToolError.VbaParseError(sourceLabel,
                        $"module '{entry.Name}' source decompression failed: {ex.Message}");
                }

                var code = cp1252.GetString(sourceBytes);
                modules.Add(new ExcelVbaModule(
                    Name: entry.Name,
                    Kind: ClassifyKind(entry, formStorageNames),
                    LineCount: CountLines(code),
                    Code: code));
            }

            return new ExcelVbaProject(HasVbaProject: true, Modules: modules);
        }
        catch (McpException) { throw; }
        catch (Exception ex)
        {
            throw ToolError.VbaParseError(sourceLabel, ex.Message);
        }
    }

    private static byte[] ReadDirStream(Storage vba, string sourceLabel)
    {
        if (!vba.TryOpenStream(DirStreamName, out var dir) || dir is null)
        {
            throw ToolError.VbaProjectLocked(sourceLabel);
        }

        using (dir)
        {
            var bytes = new byte[dir.Length];
            int read = 0;
            while (read < bytes.Length)
            {
                int n = dir.Read(bytes, read, bytes.Length - read);
                if (n == 0) break;
                read += n;
            }
            return bytes;
        }
    }

    private static string ClassifyKind(VbaModuleEntry entry, IReadOnlySet<string> formStorageNames)
    {
        if (entry.Type == 0x0021) return "standardModule";
        if (formStorageNames.Contains(entry.Name)) return "userForm";
        if (entry.Name == "ThisWorkbook" || entry.Name.StartsWith("Sheet", StringComparison.Ordinal))
            return "documentModule";
        return "classModule";
    }

    private static int CountLines(string text)
    {
        if (text.Length == 0) return 0;
        int count = 1;
        foreach (var ch in text) if (ch == '\n') count++;
        return count;
    }
}
