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
}
