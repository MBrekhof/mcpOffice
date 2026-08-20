namespace McpOffice.Models;

public sealed record PdfPageText(
    int PageNumber,
    string Text,
    int CharCount);
