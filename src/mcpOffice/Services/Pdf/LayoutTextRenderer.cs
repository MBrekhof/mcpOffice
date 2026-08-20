using McpOffice.Models;

namespace McpOffice.Services.Pdf;

/// <summary>
/// Rebuilds a page's text as a fixed-width grid, the way <c>pdftotext -layout</c> does.
///
/// DevExpress's <c>GetPageText</c> returns words in content-stream order with no horizontal
/// padding, which collapses column-based reports (lab result sheets, invoices, anything
/// originally printed as monospaced text) into ambiguous runs. Placing each word at
/// <c>round(x / charWidth)</c> keeps columns under each other, so a caller can slice the
/// output by character position.
/// </summary>
public static class LayoutTextRenderer
{
    /// <summary>Blank lines are emitted when the vertical gap exceeds this multiple of line height.</summary>
    private const double BlankLineThreshold = 1.6;

    /// <summary>Guard against a runaway blank run from a single huge gap.</summary>
    private const int MaxConsecutiveBlankLines = 4;

    public static string Render(IReadOnlyList<PdfTextLine> lines)
    {
        if (lines.Count == 0)
        {
            return string.Empty;
        }

        var charWidth = EstimateCharWidth(lines);
        if (charWidth <= 0)
        {
            return string.Join(Environment.NewLine, lines.Select(l => l.Text));
        }

        var originX = lines.SelectMany(l => l.Words).Min(w => w.X);
        var builder = new System.Text.StringBuilder();
        PdfTextLine? previous = null;

        foreach (var line in lines)
        {
            if (previous is not null)
            {
                foreach (var _ in Enumerable.Range(0, BlankLinesBetween(previous, line)))
                {
                    builder.AppendLine();
                }
            }

            builder.AppendLine(RenderLine(line, originX, charWidth));
            previous = line;
        }

        return builder.ToString().TrimEnd('\r', '\n');
    }

    private static string RenderLine(PdfTextLine line, double originX, double charWidth)
    {
        var row = new System.Text.StringBuilder();

        foreach (var word in line.Words)
        {
            var column = (int)Math.Round((word.X - originX) / charWidth);
            if (column < row.Length)
            {
                // Overlap (or a word narrower than the estimate) - never let text be lost or
                // silently merged; fall back to a single separating space.
                column = row.Length == 0 ? 0 : row.Length + 1;
            }

            row.Append(' ', column - row.Length);
            row.Append(word.Text);
        }

        return row.ToString().TrimEnd();
    }

    private static int BlankLinesBetween(PdfTextLine previous, PdfTextLine current)
    {
        var lineHeight = previous.Words.Count > 0 ? previous.Words.Max(w => w.Height) : 0;
        if (lineHeight <= 0)
        {
            return 0;
        }

        var gap = current.Y - previous.Y;
        var extra = (int)Math.Floor((gap / lineHeight) - BlankLineThreshold);
        return Math.Clamp(extra, 0, MaxConsecutiveBlankLines);
    }

    /// <summary>
    /// Median advance width per character across the page. The median (not the mean) keeps a
    /// handful of wide-glyph or single-character words from skewing the grid.
    /// </summary>
    private static double EstimateCharWidth(IReadOnlyList<PdfTextLine> lines)
    {
        var widths = lines
            .SelectMany(l => l.Words)
            .Where(w => w.Text.Length > 0 && w.Width > 0)
            .Select(w => w.Width / w.Text.Length)
            .OrderBy(w => w)
            .ToArray();

        if (widths.Length == 0)
        {
            return 0;
        }

        return widths.Length % 2 == 1
            ? widths[widths.Length / 2]
            : (widths[(widths.Length / 2) - 1] + widths[widths.Length / 2]) / 2;
    }
}
