using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using RichEditFormat = DevExpress.XtraRichEdit.DocumentFormat;

namespace McpOffice.Tests.Word;

internal static class TestWordDocuments
{
    public static string Create(Action<Document> configure) =>
        Create(server => configure(server.Document));

    public static string Create(Action<RichEditDocumentServer> configure)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.docx");

        using var server = new RichEditDocumentServer();
        configure(server);
        server.SaveDocument(path, RichEditFormat.OpenXml);

        return path;
    }

    /// <summary>
    /// Writes a fixture in a non-OpenXML format. Used for the ODT paths: DevExpress
    /// authors the .odt, so no binary fixture is committed.
    /// </summary>
    public static string CreateAs(string extension, RichEditFormat format, Action<Document> configure)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}{extension}");

        using var server = new RichEditDocumentServer();
        configure(server.Document);
        server.SaveDocument(path, format);

        return path;
    }

    public static string CreateOdt(Action<Document> configure) =>
        CreateAs(".odt", RichEditFormat.Odt, configure);

    public static void AppendParagraph(Document document, string text, string? styleName = null)
    {
        if (styleName is not null)
        {
            EnsureParagraphStyle(document, styleName);
        }

        var range = document.AppendText(text + Environment.NewLine);
        var paragraph = document.Paragraphs.Get(range).First();

        if (styleName is not null)
        {
            paragraph.Style = document.ParagraphStyles[styleName];
        }
    }

    private static void EnsureParagraphStyle(Document document, string styleName)
    {
        if (document.ParagraphStyles[styleName] is not null)
        {
            return;
        }

        var style = document.ParagraphStyles.CreateNew();
        style.Name = styleName;
        document.ParagraphStyles.Add(style);
    }
}
