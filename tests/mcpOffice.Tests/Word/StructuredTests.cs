using DevExpress.XtraRichEdit.API.Native;
using McpOffice.Models;
using McpOffice.Services.Word;

namespace McpOffice.Tests.Word;

public class StructuredTests
{
    [Fact]
    public void ReadStructured_returns_headings_paragraphs_with_bold_runs_and_tables()
    {
        var path = TestWordDocuments.Create(document =>
        {
            TestWordDocuments.AppendParagraph(document, "Title", "Heading 1");
            AppendParagraphWithBoldWord(document, "Hello ", "world", "!");
            AppendTwoByTwoTable(document, [["a1", "a2"], ["b1", "b2"]]);
        });

        var structured = new WordDocumentService().ReadStructured(path);

        var heading = Assert.IsType<HeadingBlock>(structured.Blocks[0]);
        Assert.Equal(1, heading.Level);
        Assert.Equal("Title", heading.Text);

        var paragraph = Assert.IsType<ParagraphBlock>(structured.Blocks[1]);
        Assert.Contains(paragraph.Runs, r => r.Bold && r.Text.Contains("world"));
        Assert.Contains(paragraph.Runs, r => !r.Bold && r.Text.Contains("Hello"));

        var table = Assert.Single(structured.Tables);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[0].Count);
        Assert.Equal("a1", table.Rows[0][0]);
        Assert.Equal("b2", table.Rows[1][1]);

        Assert.NotNull(structured.Properties);
    }

    private static void AppendParagraphWithBoldWord(Document document, string before, string boldText, string after)
    {
        document.AppendText(before);
        var boldRange = document.AppendText(boldText);
        var properties = document.BeginUpdateCharacters(boldRange);
        properties.Bold = true;
        document.EndUpdateCharacters(properties);
        document.AppendText(after + Environment.NewLine);
    }

    private static void AppendTwoByTwoTable(Document document, string[][] cells)
    {
        var table = document.Tables.Create(document.Range.End, cells.Length, cells[0].Length);
        for (var r = 0; r < cells.Length; r++)
        {
            for (var c = 0; c < cells[r].Length; c++)
            {
                document.InsertText(table.Rows[r].Cells[c].ContentRange.Start, cells[r][c]);
            }
        }
    }
}
