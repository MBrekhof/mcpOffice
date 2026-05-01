# mcpOffice Usage

## Requirements

- .NET 9 SDK
- DevExpress 25.2 installed locally
- DevExpress license file kept outside source control

This repo currently restores DevExpress packages from the local offline package source installed at:

```text
C:\Program Files\DevExpress 25.2\Components\System\Components\packages
```

## Build And Test

```powershell
dotnet restore
dotnet build --no-restore
dotnet test --no-restore
```

Expected current test count: 39 passing tests.

## Run The MCP Server

For local development:

```powershell
dotnet run --project C:\Projects\mcpOffice\src\mcpOffice
```

VS Code workspace config is already committed at `.vscode/mcp.json`. It starts the Debug build with:

```json
{
  "servers": {
    "office": {
      "type": "stdio",
      "command": "dotnet",
      "args": [
        "${workspaceFolder}/src/mcpOffice/bin/Debug/net9.0/mcpOffice.dll"
      ]
    }
  }
}
```

Run `dotnet build` before starting that server so the DLL exists.

For release/client configuration, publish first:

```powershell
dotnet publish C:\Projects\mcpOffice\src\mcpOffice -c Release -r win-x64 --self-contained false
```

The published executable is created under:

```text
C:\Projects\mcpOffice\src\mcpOffice\bin\Release\net9.0\win-x64\publish\mcpOffice.exe
```

Generic MCP client entry for the published executable:

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\Projects\\mcpOffice\\src\\mcpOffice\\bin\\Release\\net9.0\\win-x64\\publish\\mcpOffice.exe"
    }
  }
}
```

## Available Tools

Read tools:

- `Ping`: returns `pong`.
- `word_get_outline(path)`: returns heading nodes from a `.docx`.
- `word_get_metadata(path)`: returns core properties, page count, and word count.
- `word_read_markdown(path)`: returns a conservative Markdown projection.
- `word_read_structured(path)`: returns headings, paragraphs with runs, tables, images, and properties.
- `word_list_comments(path)`: returns comment summaries.
- `word_list_revisions(path)`: returns tracked-change summaries.

Write/create tools:

- `word_create_blank(path, overwrite=false)`: creates an empty `.docx`.
- `word_create_from_markdown(path, markdown, overwrite=false)`: creates `.docx` from Markdown.
- `word_append_markdown(path, markdown)`: appends Markdown to an existing `.docx`.
- `word_find_replace(path, find, replace, useRegex=false, matchCase=false)`: replaces text and returns replacement count.
- `word_insert_paragraph(path, atIndex, text, style?)`: inserts a paragraph.
- `word_insert_table(path, atIndex, headers[], rows[][])`: inserts a table.
- `word_set_metadata(path, properties)`: sets `author`, `title`, `subject`, and/or `keywords`.
- `word_mail_merge(templatePath, outputPath, dataJson)`: replaces `{{token}}` placeholders.

Convert tools:

- `word_convert(inputPath, outputPath, format?)`: converts to `pdf`, `html`, `rtf`, `txt`, `md`/`markdown`, or `docx`. If `format` is omitted, it is inferred from `outputPath`.

All `path`, `inputPath`, `outputPath`, and `templatePath` values must be absolute Windows paths.

## Example Calls

Create from Markdown:

```json
{
  "path": "C:\\Temp\\proposal.docx",
  "markdown": "# Proposal\n\nHello **world**.\n\n| Name | Value |\n| ---- | ----- |\n| Alpha | 1 |",
  "overwrite": false
}
```

Read outline:

```json
{
  "path": "C:\\Temp\\proposal.docx"
}
```

Expected `word_get_outline` shape:

```json
[
  { "level": 1, "text": "Proposal" }
]
```

Convert to PDF:

```json
{
  "inputPath": "C:\\Temp\\proposal.docx",
  "outputPath": "C:\\Temp\\proposal.pdf"
}
```

Mail merge:

```json
{
  "templatePath": "C:\\Temp\\template.docx",
  "outputPath": "C:\\Temp\\merged.docx",
  "dataJson": "{\"firstName\":\"Ada\",\"score\":42}"
}
```

Expected `word_get_metadata` shape:

```json
{
  "author": "Martin",
  "title": "Proposal",
  "subject": "MCP Office",
  "keywords": "mcp,office,word",
  "created": "2026-04-30T10:00:00",
  "modified": "2026-04-30T11:00:00",
  "lastPrinted": null,
  "revisionCount": 7,
  "pageCount": 1,
  "wordCount": 1200
}
```

## Markdown Notes

`word_create_from_markdown` and `word_append_markdown` use `MarkdownToDocxGenerator` for richer Markdown import. Current coverage includes headings, paragraphs, bold, common italic spans, simple lists, fenced code blocks, links/images at the package level, and tables.

Known caveats:

- Lists currently round-trip through `word_read_structured` as paragraph text with literal `-` or `1.` prefixes, not semantic Word list objects.
- Hyperlink URLs are not exposed by `word_read_structured` yet.
- Markdown export (`word_read_markdown` and `.md` conversion) is a conservative projection, not a full-fidelity Markdown serializer.

## Error Codes

Tool errors are returned as `McpException` messages prefixed with stable codes:

- `[file_not_found]`
- `[file_exists]`
- `[invalid_path]`
- `[unsupported_format]`
- `[parse_error]`
- `[index_out_of_range]`
- `[merge_field_missing]`
- `[io_error]`
- `[internal_error]`

## Troubleshooting

- If restore cannot find DevExpress packages, confirm DevExpress 25.2 is installed or update `nuget.config` to point at your installed offline package path.
- If VS Code cannot start the MCP server, run `dotnet build` and confirm `src\mcpOffice\bin\Debug\net9.0\mcpOffice.dll` exists.
- If tool calls fail with `[invalid_path]`, pass an absolute path such as `C:\Docs\file.docx`.
- If tool calls fail with `[file_not_found]`, confirm the MCP server process can access the file.
- If output calls fail with `[file_exists]`, use a different output path or pass `overwrite=true` where the tool supports it.
- If the MCP client hangs, verify the server logs go to stderr only; stdout is reserved for the MCP JSON-RPC stream.
