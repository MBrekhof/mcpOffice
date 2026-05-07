# mcpOffice — `excel_export_csv` Design

**Date:** 2026-05-07
**Status:** Approved (brainstorming phase)
**Scope:** New MCP tool that streams a worksheet (or A1 range) to a CSV file on disk for `pandas.read_csv` / `polars.read_csv` consumption. Replaces the JSON cell-grid path for "load this sheet as a dataframe" workflows.

## Purpose

`excel_read_sheet` is the right shape for surgical reads — slice an A1 range, get back a JSON cell-grid with per-cell type metadata. It's the wrong shape for "dump this whole sheet, I'll process it in pandas." The 50,000-cell cap rejects most real datasets, and reassembling a 100k-row JSON page-by-page is wasted tokens and wasted time.

`excel_export_csv` writes the sheet straight to a CSV file the agent already knows how to consume. The deliverable is a path on disk, not a JSON payload — so it's bounded by file size, not LLM context.

Target consumer: agents loading Excel data into a Python data-tools workflow (`pandas`, `polars`, `duckdb`).

## Operation model

Same stateless / file-path shape as the rest of the Excel surface. Internally reuses `LoadWorkbook`, `ResolveWorksheet`, and `GetCellValue` from `ExcelWorkbookService` — no new workbook plumbing. The new code is a streaming CSV writer that handles RFC 4180 quoting and invariant-culture formatting.

Every call is a one-shot: open workbook → resolve sheet → walk range → write file → close everything. No formula recalculation; cells are read at their cached values (whatever Excel had on last save). Stale cached values are out of scope — the agent uses `excel_list_formulas` if it needs recalc semantics.

## Tool surface

```
excel_export_csv(
    path: string,                 // absolute path to .xlsx/.xlsm input
    outputPath: string,           // absolute path to .csv output
    sheetName?: string,           // mirrors excel_read_sheet
    sheetIndex?: int,             //   "
    range?: string,               // A1 range; defaults to worksheet's used range
    overwrite: bool = false,      // matches word_create_blank — false errors if outputPath exists
    maxRows: int = 1_048_576      // Excel's row ceiling; safety guard against pathological input
) -> ExcelExportCsvResult
```

This joins as the **27th** MCP tool. The existing 26 tools stay untouched.

### Sheet selection

Mirrors `excel_read_sheet`'s `sheetName` / `sheetIndex` / `range` triple verbatim:
- Both omitted → sheet at index 0.
- `sheetName` set → resolve by name (case-sensitive, matches DevExpress behaviour).
- `sheetIndex` set → resolve by 0-based index.
- Both set → `sheetName` wins (matches existing `ResolveWorksheet` precedence).
- `range` null/empty → uses `worksheet.GetUsedRange()`.

## Output schema

```jsonc
{
  "outputPath": "C:\\path\\to\\output.csv",
  "rowCount": 12345,             // rows actually written (== range.RowCount)
  "columnCount": 42,             // columns actually written (== range.ColumnCount)
  "bytesWritten": 8732145         // size of the written file in bytes
}
```

Returning a structured result (rather than just `outputPath`) lets the agent confirm dimensions without a follow-up `stat` call, and makes the result visually distinguishable from `word_*` tools that return the path verbatim.

## CSV dialect

| Aspect | Choice | Rationale |
|---|---|---|
| Encoding | UTF-8, **no BOM** | `pandas.read_csv` default; BOM breaks naive consumers. |
| Delimiter | Comma `,` | RFC 4180; pandas/polars default. |
| Line ending | CRLF `\r\n` | RFC 4180; both pandas and polars normalise on read. |
| Text quoting | RFC 4180 minimal: wrap in `"…"` only when value contains `,`, `"`, `\r`, or `\n`; embedded `"` doubled to `""` | Smaller files, identical parse behaviour. |
| Numbers | Invariant culture, no thousand separator (`0.21`, `1234567.89`) | Locale-neutral; `pd.read_csv` parses without `thousands=` hint. |
| DateTime | ISO 8601 with `T` separator and seconds (`2026-05-07T14:30:00`); midnight emits `T00:00:00` | Unambiguous; `pd.read_csv(parse_dates=...)` handles natively. |
| Boolean | `true` / `false` (lowercase) | Pandas-friendly. Excel surfaces booleans as `bool`, not 0/1. |
| Empty cell | Empty field (no quoting) | Pandas reads as `NaN` (default `na_values`). |
| Formula cell | Cached value via the same value/type extraction as `excel_read_sheet` (no formula text) | "Values, not formulas" — the whole point of the tool. |

No `na_rep`, no decimal/thousand-separator parameter, no header-row option. These are configurable on the consumer side (`pandas.read_csv` covers them) and adding them on the producer side multiplies the test surface for no real gain.

## Error handling

Reuses the existing error code set; no new codes needed.

| Code | When |
|---|---|
| `file_not_found` | input `path` missing — `PathGuard.RequireExists`. |
| `invalid_path` | `path` or `outputPath` non-absolute / empty — `PathGuard.RequireAbsolute` / `RequireWritable`. |
| `file_exists` | `outputPath` exists and `overwrite=false` — `PathGuard.RequireWritable`. |
| `index_out_of_range` | `sheetIndex` outside `[0, sheetCount-1]` — existing `ResolveWorksheet`. |
| `parse_error` | DevExpress fails to load workbook — existing wrapper in service. |
| `range_too_large` | `range.RowCount > maxRows` — extended message: `"sheet 'Foo' rows X exceeds maxRows Y"`. |
| `io_error` | disk write fails (out of space, locked file, permissions) — surfaces OS message. |

## Architecture

```
Tools/ExcelTools.cs
  ExcelExportCsv(...)                       -- one-line delegate
        |
        v
Services/Excel/ExcelWorkbookService.cs
  ExportCsv(path, outputPath, sheetName,
            sheetIndex, range, overwrite,
            maxRows) : ExcelExportCsvResult -- orchestrator
        |     |
        |     +-- LoadWorkbook (existing)
        |     +-- ResolveWorksheet (existing)
        |     +-- GetCellValue / GetCellValueType (existing)
        v
Services/Excel/Csv/CsvWriter.cs             -- new: streaming writer + RFC 4180 quoting
  WriteRow(IReadOnlyList<object?>)
  Flush() / Dispose()
```

Why a separate `CsvWriter` rather than inlining: the quoting + invariant-culture formatting is small (~80 LOC) but has a wide test surface (every cell type × every quoting trigger). Splitting it out keeps the unit tests focused on string output, decoupled from workbook fixtures.

`ExcelExportCsvResult` is a new record under `Models/`; same shape as `ReplaceResult` etc.

## Testing

### Unit (`tests/mcpOffice.Tests/Excel/`)
- `CsvWriterTests` — pure string output, no workbook involved:
  - Number, datetime, bool, text, null formatting (round-trips through invariant culture).
  - Quoting triggers: comma, double-quote, CR, LF, mix; embedded `"` doubled.
  - Multiple rows, mixed types per row.
  - CRLF line endings between rows; no trailing CRLF after final row (pandas-friendly).
  - UTF-8 without BOM (assert raw byte prefix).
- `ExcelExportCsvTests` — service against programmatic fixtures (extends `TestExcelWorkbooks`):
  - Happy path: 5 rows × 4 cols mixed types → file matches expected CSV byte-for-byte.
  - `range` slicing: workbook has 100 rows, `range="A1:B10"` writes 10 rows × 2 cols.
  - `sheetName` and `sheetIndex` resolution; precedence when both set.
  - `overwrite=false` + existing file → `file_exists`.
  - `overwrite=true` + existing file → succeeds, file is replaced.
  - Output directory created if missing.
  - `maxRows=5` against a 10-row sheet → `range_too_large`.
  - Formula cell writes the cached value, not formula text.
  - Empty cells emit empty fields (no `null`, no quotes).
  - Invariant-culture confirmation: a workbook authored with `nl-NL` host culture still emits `0.21` (not `0,21`).
  - Result record's `rowCount`, `columnCount`, `bytesWritten` are correct.

### Integration (`tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs`)
- One end-to-end happy path through stdio: spawn server, call `excel_export_csv`, read the resulting file, assert first line is the expected CSV header row.
- Tool catalog: `ToolSurfaceTests` adds `excel_export_csv` to the expected list (now 27 names).

### Live verification
- One run against `C:\Projects\mcpOffice-samples\Air.xlsm` exporting one sheet to a temp CSV; load with `pandas.read_csv` in a quick PowerShell + `python -c` check, assert dataframe shape matches `excel_list_sheets`'s reported `rowCount` × `columnCount`. Same gate as the existing Air.xlsm benchmark — skips when the file is absent.

## What this design deliberately does NOT do

- **No NDJSON sibling.** `excel_export_ndjson` is on TODO; it shares streaming infrastructure with this tool but answers a different question (column-typed output for `pandas.read_json(lines=True)`). Ship CSV first; NDJSON lands on top.
- **No `.csv.gz` compression.** Trivial follow-up — wrap the `FileStream` in `GZipStream` when `outputPath` ends in `.gz`. Adds a test matrix that's not justified for v1.
- **No header parameter.** Consumer-side concern (`pandas.read_csv(header=0)` is the convention).
- **No formula-text mode.** That's `excel_list_formulas`'s job.
- **No multi-sheet export.** A CSV file is a single grid by definition. Multi-sheet means writing N files, which is the agent's call (loop + N invocations).
- **No formula recalculation.** Cached value only. `excel_list_formulas` covers the recalc path.
- **No pivot / chart export.** CSV is row-and-column data; structural Excel features don't have a sensible CSV projection.

## Open follow-ups (logged on TODO, not addressed here)

- Compression via `.csv.gz` extension sniff.
- `excel_export_ndjson` sibling.
- Optional `lineEnding` parameter (`crlf` / `lf`) — only if a consumer surfaces a real need.
- Optional `delimiter` parameter (`,` / `\t` / `;`) — same gate.
