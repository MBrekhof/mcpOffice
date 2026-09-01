namespace McpOffice.Tests.Integration;

public class ToolSurfaceTests
{
    [Fact]
    public async Task Exposes_initial_tool_catalog()
    {
        string[] expected =
        [
            "excel_analyze_vba",
            "excel_export_csv",
            "excel_extract_vba",
            "excel_get_metadata",
            "excel_get_structure",
            "excel_list_defined_names",
            "excel_list_formulas",
            "excel_list_sheets",
            "excel_list_vba_entry_points",
            "excel_read_sheet",
            "excel_render_vba_callgraph",
            "excel_suggest_vba_conversion",
            "pdf_extract_images",
            "pdf_find_text",
            "pdf_get_metadata",
            "pdf_get_outline",
            "pdf_read_layout",
            "pdf_read_text",
            "pdf_render_page",
            "Ping",
            "word_append_markdown",
            "word_convert",
            "word_create_blank",
            "word_create_from_markdown",
            "word_find_replace",
            "word_get_metadata",
            "word_get_outline",
            "word_insert_paragraph",
            "word_insert_table",
            "word_list_comments",
            "word_list_revisions",
            "word_mail_merge",
            "word_read_markdown",
            "word_read_structured",
            "word_set_metadata"
        ];

        await using var harness = await ServerHarness.StartAsync();
        var tools = await harness.Client.ListToolsAsync();
        var toolNames = tools.Select(t => t.Name).ToHashSet();

        Assert.Equal(expected.Order(), toolNames.Order());
    }
}
