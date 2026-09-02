using RichEditFormat = DevExpress.XtraRichEdit.DocumentFormat;

namespace McpOffice.Services.Word;

/// <summary>
/// Maps a file extension to the RichEdit format to load/save it with.
/// <para>
/// RichEditDocumentServer does not sniff the format when one is passed explicitly, and
/// passing OpenXml for an .odt fails the load outright. Every Word load and in-place
/// save goes through here so the two stay in agreement: a file is written back in the
/// format it was read as.
/// </para>
/// <para>
/// OpenXml is the fallback for an unknown or missing extension — that is the historical
/// behaviour and .docx remains the common case.
/// </para>
/// </summary>
internal static class WordFormats
{
    public static RichEditFormat ForPath(string path) =>
        Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".odt" => RichEditFormat.Odt,
            ".rtf" => RichEditFormat.Rtf,
            ".doc" => RichEditFormat.Doc,
            ".dot" => RichEditFormat.Dot,
            ".docm" => RichEditFormat.Docm,
            ".dotx" => RichEditFormat.Dotx,
            ".dotm" => RichEditFormat.Dotm,
            ".txt" => RichEditFormat.PlainText,
            ".htm" or ".html" => RichEditFormat.Html,
            ".mht" => RichEditFormat.Mht,
            ".epub" => RichEditFormat.ePub,
            ".xml" => RichEditFormat.WordML,
            _ => RichEditFormat.OpenXml
        };
}
