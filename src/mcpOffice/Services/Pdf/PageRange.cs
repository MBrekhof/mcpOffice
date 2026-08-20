namespace McpOffice.Services.Pdf;

/// <summary>
/// Parses the <c>pageRange</c> parameter shared by the PDF tools: a comma-separated list of
/// 1-based page numbers and inclusive spans, e.g. <c>"1"</c>, <c>"2-5"</c>, <c>"1,3,7-9"</c>.
/// An open-ended span <c>"5-"</c> runs to the last page. Null/blank means every page.
/// </summary>
public static class PageRange
{
    public static IReadOnlyList<int> Parse(string? range, int pageCount)
    {
        if (pageCount <= 0)
        {
            return [];
        }

        if (string.IsNullOrWhiteSpace(range))
        {
            return Enumerable.Range(1, pageCount).ToArray();
        }

        var pages = new SortedSet<int>();

        foreach (var rawToken in range.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            var token = rawToken.Trim();
            if (token.Length == 0)
            {
                continue;
            }

            var dash = token.IndexOf('-');
            if (dash < 0)
            {
                pages.Add(ParsePage(token, range, pageCount));
                continue;
            }

            var leftText = token[..dash].Trim();
            var rightText = token[(dash + 1)..].Trim();

            if (leftText.Length == 0)
            {
                throw ToolError.InvalidPageRange(range, $"span '{token}' has no start page.");
            }

            var from = ParsePage(leftText, range, pageCount);
            // "5-" means "from 5 to the end".
            var to = rightText.Length == 0 ? pageCount : ParsePage(rightText, range, pageCount);

            if (to < from)
            {
                throw ToolError.InvalidPageRange(range, $"span '{token}' ends before it starts.");
            }

            for (var page = from; page <= to; page++)
            {
                pages.Add(page);
            }
        }

        if (pages.Count == 0)
        {
            throw ToolError.InvalidPageRange(range, "no pages were selected.");
        }

        return pages.ToArray();
    }

    private static int ParsePage(string text, string range, int pageCount)
    {
        if (!int.TryParse(text, out var page))
        {
            throw ToolError.InvalidPageRange(range, $"'{text}' is not a number.");
        }

        if (page < 1 || page > pageCount)
        {
            throw ToolError.PageNotFound(page, pageCount);
        }

        return page;
    }
}
