using McpOffice.Models;

namespace McpOffice.Services.Pdf;

/// <summary>
/// Turns a bag of positioned words into visual lines.
///
/// PDF has no concept of a line of text — a page is a set of glyph-drawing operations — so
/// "which words share a line" is a clustering decision, not something the format records.
/// Words are grouped when their vertical centres fall within a tolerance derived from word
/// height, which tracks font size instead of assuming one leading for the whole document.
/// </summary>
public static class LineGrouper
{
    /// <summary>Fraction of word height two words may differ by and still count as one line.</summary>
    private const double DefaultTolerance = 0.6;

    public static IReadOnlyList<PdfTextLine> Group(
        int pageNumber,
        IReadOnlyList<PdfWordBox> words,
        double toleranceFactor = DefaultTolerance)
    {
        if (words.Count == 0)
        {
            return [];
        }

        // Reading order: top of the page first (Y already grows downward), then left to right.
        var ordered = words.OrderBy(w => w.Y).ThenBy(w => w.X).ToList();

        var lines = new List<List<PdfWordBox>>();
        var current = new List<PdfWordBox> { ordered[0] };
        var currentCentre = Centre(ordered[0]);

        for (var i = 1; i < ordered.Count; i++)
        {
            var word = ordered[i];
            // Use the taller of the two words so a small superscript next to body text still
            // joins the line it visually belongs to.
            var tolerance = Math.Max(word.Height, current[^1].Height) * toleranceFactor;

            if (Math.Abs(Centre(word) - currentCentre) <= tolerance)
            {
                current.Add(word);
                // Track the running centre so a line that drifts slightly does not split.
                currentCentre = current.Average(Centre);
            }
            else
            {
                lines.Add(current);
                current = [word];
                currentCentre = Centre(word);
            }
        }

        lines.Add(current);

        return lines
            .Select(line =>
            {
                var byX = line.OrderBy(w => w.X).ToArray();
                return new PdfTextLine(
                    pageNumber,
                    byX[0].X,
                    byX.Min(w => w.Y),
                    string.Join(' ', byX.Select(w => w.Text)),
                    byX);
            })
            .ToArray();
    }

    private static double Centre(PdfWordBox word) => word.Y + (word.Height / 2);
}
