# mcpOffice

An MCP (Model Context Protocol) server for Microsoft Office documents, written in C# (.NET 9) and backed by DevExpress Office File API packages. Built so AI agents (Claude Code, Claude Desktop, etc.) can read, write, and convert Office documents through tool calls instead of one-off scripts.

**Status:** Word (.docx) POC in progress. The server currently exposes outline, metadata, and Markdown-read tools.

## Documents

- [Usage](docs/usage.md) - build, run, MCP client config, and sample calls.
- [Design](docs/plans/2026-04-30-mcpoffice-word-poc-design.md) - architecture, tool surface, error model, packaging.
- [Implementation plan](docs/plans/2026-04-30-mcpoffice-word-poc-plan.md) - task-by-task TDD plan.

## Current Tools

- `Ping`
- `word_get_outline(path)`
- `word_get_metadata(path)`
- `word_read_markdown(path)`

## Roadmap

1. **Word POC** - read / write / convert .docx (current).
2. Excel (.xlsx).
3. PowerPoint (.pptx).
4. PDF.

## Built with

- [`ModelContextProtocol`](https://github.com/modelcontextprotocol/csharp-sdk) - official C# MCP SDK.
- DevExpress RichEdit / Office File API packages - server-side Word document APIs (requires a DevExpress license).
