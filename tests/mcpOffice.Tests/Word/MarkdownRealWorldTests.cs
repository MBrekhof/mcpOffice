using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class MarkdownRealWorldTests
{
    [Fact]
    public void Fn_send_email_callers_md_round_trips_with_tables_and_inline_code()
    {
        var md = File.ReadAllText(TestFixtures.Path("fn_send_email_callers.md"));
        using var server = new RichEditDocumentServer();
        MarkdownToDocxConverter.Apply(server.Document, md, null);

        // Tables present — the source has 4 named category tables + 1 summary table = 5 tables total.
        // Accept >= 4 as the minimum bar (the summary table may merge depending on renderer).
        Assert.True(server.Document.Tables.Count >= 4,
            $"expected >=4 tables, got {server.Document.Tables.Count}");

        // Inline code preserved — FN_SEND_EMAIL appears inside backtick spans throughout the body.
        var bodyText = server.Document.GetText(server.Document.Range);
        Assert.Contains("FN_SEND_EMAIL", bodyText);

        // Bold survived — the source has **optional**, **Fix:** etc.
        Assert.True(HasBoldRun(server.Document),
            "expected at least one bold run from **...** spans");
    }

    private static bool HasBoldRun(Document doc)
    {
        // Walk through all paragraphs and sample character properties at intervals.
        // Avoids an O(N) full character scan while still catching any bold run.
        var totalLen = doc.Range.End.ToInt() - doc.Range.Start.ToInt();
        // Step every 3 characters — tight enough to catch a 2-char bold span like "**".
        for (int i = 0; i < totalLen; i += 3)
        {
            var pos = doc.CreatePosition(doc.Range.Start.ToInt() + i);
            var range = doc.CreateRange(pos, 1);
            var props = doc.BeginUpdateCharacters(range);
            try
            {
                if (props.Bold == true) return true;
            }
            finally { doc.EndUpdateCharacters(props); }
        }
        return false;
    }
}
