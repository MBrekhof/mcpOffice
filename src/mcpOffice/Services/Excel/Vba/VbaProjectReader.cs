using System.IO.Compression;
using System.Text;
using System.Xml;
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

            var codenames = ExtractDocumentModuleCodenames(zip);

            using var ms = new MemoryStream();
            using (var s = entry.Open()) s.CopyTo(ms);
            ms.Position = 0;
            return ReadVbaProjectBin(ms, xlsmPath, codenames);
        }
        catch (McpException) { throw; }
        catch (Exception ex)
        {
            throw ToolError.VbaParseError(xlsmPath, ex.Message);
        }
    }

    public ExcelVbaProject ReadVbaProjectBin(
        Stream vbaProjectBin,
        string sourceLabel,
        IReadOnlySet<string>? documentModuleCodenames = null)
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

            // UserForms have a same-named sub-storage at the root level (next to VBA),
            // holding layout streams like f/o. Excel writes them at root only; the
            // MS-OVBA spec doesn't define sub-storages inside VBA, so we don't scan there.
            var formStorageNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var info in root.EnumerateEntries())
            {
                if (info.Type == EntryType.Storage &&
                    !string.Equals(info.Name, VbaStorageName, StringComparison.OrdinalIgnoreCase))
                {
                    formStorageNames.Add(info.Name);
                }
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
                    Kind: ClassifyKind(entry, formStorageNames, documentModuleCodenames),
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

    private static string ClassifyKind(
        VbaModuleEntry entry,
        IReadOnlySet<string> formStorageNames,
        IReadOnlySet<string>? documentModuleCodenames)
    {
        if (entry.Type == 0x0021) return "standardModule";
        if (formStorageNames.Contains(entry.Name)) return "userForm";

        // Primary signal: the module name appears as a codename in the OOXML
        // (workbookPr/codeName or any sheet's sheetPr/codeName). Locale-independent
        // and survives user-renamed codenames.
        if (documentModuleCodenames is not null)
        {
            return documentModuleCodenames.Contains(entry.Name) ? "documentModule" : "classModule";
        }

        // Fallback for callers without OOXML context (synthetic VbaProjectBinBuilder
        // tests, etc.): English-default codename heuristic. Will misclassify Dutch
        // (Blad), German (Tabelle), Italian (Foglio), French (Feuil), Spanish (Hoja)
        // sheet codenames as classModule — pass documentModuleCodenames to avoid this.
        if (entry.Name == "ThisWorkbook" || entry.Name.StartsWith("Sheet", StringComparison.Ordinal))
            return "documentModule";
        return "classModule";
    }

    /// <summary>
    /// Builds the set of codenames used by document modules in the workbook by
    /// scanning the OOXML manifest and per-sheet xmls. Best-effort: returns whatever
    /// could be read; ThisWorkbook is always added as a baseline because some old
    /// workbooks omit workbookPr/codeName even though the module is always present.
    /// </summary>
    private static IReadOnlySet<string> ExtractDocumentModuleCodenames(ZipArchive zip)
    {
        var codenames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ThisWorkbook" };

        // Workbook codename — typically "ThisWorkbook" but can be customized.
        TryAddCodeName(zip, "xl/workbook.xml", "workbookPr", codenames);

        // Sheet codenames — scan worksheets, chartsheets, dialogsheets.
        foreach (var entry in zip.Entries)
        {
            if (IsSheetXmlEntry(entry.FullName))
            {
                TryAddCodeName(zip, entry.FullName, "sheetPr", codenames);
            }
        }

        return codenames;
    }

    private static bool IsSheetXmlEntry(string fullName)
    {
        if (!fullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return false;
        return fullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
            || fullName.StartsWith("xl/chartsheets/", StringComparison.OrdinalIgnoreCase)
            || fullName.StartsWith("xl/dialogsheets/", StringComparison.OrdinalIgnoreCase);
    }

    private static void TryAddCodeName(
        ZipArchive zip, string entryName, string elementLocalName, HashSet<string> codenames)
    {
        var entry = zip.GetEntry(entryName);
        if (entry is null) return;
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
            });
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (!string.Equals(reader.LocalName, elementLocalName, StringComparison.Ordinal)) continue;
                var cn = reader.GetAttribute("codeName");
                if (!string.IsNullOrEmpty(cn)) codenames.Add(cn);
                return; // first match is the only one we want for this file
            }
        }
        catch
        {
            // Best-effort: codename extraction enhances the classifier; failure here
            // falls back to the legacy heuristic at the ClassifyKind site. No throw.
        }
    }

    private static int CountLines(string text)
    {
        if (text.Length == 0) return 0;
        int count = 1;
        foreach (var ch in text) if (ch == '\n') count++;
        return count;
    }
}
