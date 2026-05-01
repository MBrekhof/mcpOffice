using McpOffice.Models;

namespace McpOffice.Services.Word;

public interface IWordDocumentService
{
    IReadOnlyList<OutlineNode> GetOutline(string path);
    DocumentMetadata GetMetadata(string path);
    string ReadAsMarkdown(string path);
    StructuredDocument ReadStructured(string path);
    IReadOnlyList<CommentEntry> ListComments(string path);
    IReadOnlyList<RevisionEntry> ListRevisions(string path);
    string CreateBlank(string path, bool overwrite);
    string CreateFromMarkdown(string path, string markdown, bool overwrite);
    string AppendMarkdown(string path, string markdown);
    ReplaceResult FindReplace(string path, string find, string replace, bool useRegex, bool matchCase);
    string InsertParagraph(string path, int atIndex, string text, string? style);
    string InsertTable(string path, int atIndex, IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows);
    string SetMetadata(string path, IReadOnlyDictionary<string, string> properties);
    string MailMerge(string templatePath, string outputPath, string dataJson);
    string Convert(string inputPath, string outputPath, string? format);
}
