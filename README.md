# mcpOffice

An MCP (Model Context Protocol) server for Microsoft Office documents, written in C# (.NET 9) and backed by [DevExpress.Docs](https://www.devexpress.com/products/net/office-file-api/). Built so AI agents (Claude Code, Claude Desktop, etc.) can read, write, and convert Office documents through tool calls instead of one-off Python scripts.

**Status:** Planning complete. Word (.docx) POC implementation pending.

## Documents

- [Design](docs/plans/2026-04-30-mcpoffice-word-poc-design.md) — architecture, tool surface, error model, packaging.
- [Implementation plan](docs/plans/2026-04-30-mcpoffice-word-poc-plan.md) — task-by-task TDD plan.

## Roadmap

1. **Word POC** — read / write / convert .docx (current).
2. Excel (.xlsx).
3. PowerPoint (.pptx).
4. PDF.

## Built with

- [`ModelContextProtocol`](https://github.com/modelcontextprotocol/csharp-sdk) — official C# MCP SDK.
- [`DevExpress.Docs`](https://www.devexpress.com/products/net/office-file-api/) — server-side Office file APIs (requires a DevExpress license).
