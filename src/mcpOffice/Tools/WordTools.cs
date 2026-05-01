using System.ComponentModel;
using McpOffice.Services.Word;
using ModelContextProtocol.Server;

namespace McpOffice.Tools;

[McpServerToolType]
public static class WordTools
{
    private static readonly IWordDocumentService Service = new WordDocumentService();

    [McpServerTool(Name = "word_get_outline")]
    [Description("Returns the heading tree of a .docx file as [{level,text}]. Cheap; use to skim structure.")]
    public static object WordGetOutline(
        [Description("Absolute path to the .docx file")] string path)
        => Service.GetOutline(path);

    [McpServerTool(Name = "word_get_metadata")]
    [Description("Returns core .docx metadata, page count, and word count.")]
    public static object WordGetMetadata(
        [Description("Absolute path to the .docx file")] string path)
        => Service.GetMetadata(path);

    [McpServerTool(Name = "word_read_markdown")]
    [Description("Returns a Markdown projection of a .docx file. Preserves headings and paragraph text.")]
    public static string WordReadMarkdown(
        [Description("Absolute path to the .docx file")] string path)
        => Service.ReadAsMarkdown(path);

    [McpServerTool(Name = "word_read_structured")]
    [Description("Returns a structured tree of blocks (headings, paragraphs with runs), tables, images, and document properties. Use for surgical edits or when run-level formatting matters.")]
    public static object WordReadStructured(
        [Description("Absolute path to the .docx file")] string path)
        => Service.ReadStructured(path);

    [McpServerTool(Name = "word_list_comments")]
    [Description("Returns all comments in a .docx file: id, author, date, comment body text, and the anchored text it relates to.")]
    public static object WordListComments(
        [Description("Absolute path to the .docx file")] string path)
        => Service.ListComments(path);

    [McpServerTool(Name = "word_list_revisions")]
    [Description("Returns all tracked-change revisions: type (insert/delete/format/...), author, date, and affected text.")]
    public static object WordListRevisions(
        [Description("Absolute path to the .docx file")] string path)
        => Service.ListRevisions(path);
}
