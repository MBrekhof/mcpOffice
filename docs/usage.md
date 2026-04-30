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

## Run The MCP Server

For local development:

```powershell
dotnet run --project C:\Projects\mcpOffice\src\mcpOffice
```

For client configuration, publish first:

```powershell
dotnet publish C:\Projects\mcpOffice\src\mcpOffice -c Release -r win-x64 --self-contained false
```

The published executable is created under:

```text
C:\Projects\mcpOffice\src\mcpOffice\bin\Release\net9.0\win-x64\publish\mcpOffice.exe
```

## MCP Client Config

Example MCP server entry:

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

- `Ping`: returns `pong`.
- `word_get_outline(path)`: returns heading nodes from a `.docx`.
- `word_get_metadata(path)`: returns core properties, page count, and word count.
- `word_read_markdown(path)`: returns a conservative Markdown projection of headings and paragraph text.

All `path` values must be absolute Windows paths.

## Example Calls

```json
{
  "path": "C:\\Users\\you\\Documents\\proposal.docx"
}
```

Expected `word_get_outline` shape:

```json
[
  { "level": 1, "text": "Introduction" },
  { "level": 2, "text": "Background" }
]
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

## Troubleshooting

- If restore cannot find DevExpress packages, confirm DevExpress 25.2 is installed or update `nuget.config` to point at your installed offline package path.
- If tool calls fail with `[invalid_path]`, pass an absolute path such as `C:\Docs\file.docx`.
- If tool calls fail with `[file_not_found]`, confirm the MCP server process can access the file.
- If the MCP client hangs, verify the server logs go to stderr only; stdout is reserved for the MCP JSON-RPC stream.
