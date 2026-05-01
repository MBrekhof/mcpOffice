# mcpOffice

An MCP (Model Context Protocol) server for Microsoft Office documents, written in C# (.NET 9) and backed by DevExpress Office File API packages. It lets AI agents read, write, and convert Office documents through tool calls instead of one-off scripts.

**Status:** Word (.docx) POC is feature-complete through read/write/convert tools. Final release verification is next.

## Documents

- [Usage](docs/usage.md) - build, run, MCP client config, sample calls, troubleshooting.
- [Design](docs/plans/2026-04-30-mcpoffice-word-poc-design.md) - architecture, tool surface, error model, packaging.
- [Implementation plan](docs/plans/2026-04-30-mcpoffice-word-poc-plan.md) - task-by-task TDD plan.

## Current Tools

- `Ping`
- `word_get_outline(path)`
- `word_get_metadata(path)`
- `word_read_markdown(path)`
- `word_read_structured(path)`
- `word_list_comments(path)`
- `word_list_revisions(path)`
- `word_create_blank(path, overwrite=false)`
- `word_create_from_markdown(path, markdown, overwrite=false)`
- `word_append_markdown(path, markdown)`
- `word_find_replace(path, find, replace, useRegex=false, matchCase=false)`
- `word_insert_paragraph(path, atIndex, text, style?)`
- `word_insert_table(path, atIndex, headers[], rows[][])`
- `word_set_metadata(path, properties)`
- `word_mail_merge(templatePath, outputPath, dataJson)`
- `word_convert(inputPath, outputPath, format?)`

All file paths passed to tools must be absolute.

## Example

Create a Word document from Markdown, then convert it to PDF:

```json
{
  "path": "C:\\Temp\\proposal.docx",
  "markdown": "# Proposal\n\nHello **Word**.",
  "overwrite": false
}
```

```json
{
  "inputPath": "C:\\Temp\\proposal.docx",
  "outputPath": "C:\\Temp\\proposal.pdf"
}
```

## Roadmap

1. **Word POC** - read / write / convert .docx (current).
2. Excel (.xlsx).
3. PowerPoint (.pptx).
4. PDF.

## Built With

- [`ModelContextProtocol`](https://github.com/modelcontextprotocol/csharp-sdk) - C# MCP SDK.
- DevExpress RichEdit / Office File API packages - server-side Word document APIs.
- [`MarkdownToDocxGenerator`](https://www.nuget.org/packages/MarkdownToDocxGenerator) - richer Markdown-to-DOCX import.
