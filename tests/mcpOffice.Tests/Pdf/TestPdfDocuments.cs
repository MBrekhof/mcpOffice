using DevExpress.XtraRichEdit;
using RichEditFormat = DevExpress.XtraRichEdit.DocumentFormat;

namespace McpOffice.Tests.Pdf;

/// <summary>
/// Generates PDF fixtures programmatically, per the repo convention of not committing binary
/// fixtures. Built by rendering a RichEdit document to PDF - the same DevExpress export path
/// word_convert already uses - so the fixtures need no PDF-authoring API of their own.
/// </summary>
internal static class TestPdfDocuments
{
    /// <summary>A PDF with the given paragraphs, one per line, in document order.</summary>
    public static string Create(params string[] paragraphs)
        => Create(configure: null, paragraphs);

    public static string Create(Action<RichEditDocumentServer>? configure, params string[] paragraphs)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.pdf");

        using var server = new RichEditDocumentServer();
        foreach (var paragraph in paragraphs)
        {
            server.Document.AppendText(paragraph + Environment.NewLine);
        }

        configure?.Invoke(server);
        server.ExportToPdf(path);

        return path;
    }

    /// <summary>A multi-page PDF: <paramref name="pageCount"/> pages, each holding a marker line.</summary>
    public static string CreateMultiPage(int pageCount)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcpoffice-{Guid.NewGuid():N}.pdf");

        using var server = new RichEditDocumentServer();
        var document = server.Document;

        for (var page = 1; page <= pageCount; page++)
        {
            document.AppendText($"Page marker {page}" + Environment.NewLine);
            if (page < pageCount)
            {
                document.AppendSection();
            }
        }

        server.ExportToPdf(path);
        return path;
    }

    /// <summary>Deletes a fixture, ignoring a file that has already gone.</summary>
    public static void Delete(params string[] paths)
    {
        foreach (var path in paths)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { /* a locked temp file is not a test failure */ }
        }
    }
}
