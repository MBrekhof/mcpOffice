# mcpOffice — Excel write tools (create / new version) — Design

**Date:** 2026-11-04
**Card:** ContextBoard #2853 "Create Excel files" (project mcpOffice, id 27)
**Status:** Draft, awaiting approval

## Why

A Claude session regularly ends up holding data a person wants as a spreadsheet: query results
from LabWare PRD, a comparison, a list for a colleague. Today the session writes a CSV or a
throwaway Python script, which the repo rule says means a tool is missing. The session should be
able to hand over a proper `.xlsx` in one call: readable at a glance, numbers stored as numbers,
dates as dates. The second case is taking an existing workbook and producing a **new version**
of it.

## Goals

- Create a formatted workbook from tabular data in **one call** for the common case.
- Produce a new version of an existing `.xlsx` / `.xlsm`: write values and formulas, format,
  add, rename and delete sheets, insert and delete rows and columns. All of that happens in **one
  call** that saves under a new name.
- Output that reads well without further work: bold frozen header, filter, sensible column widths.
- One shared vocabulary of operations for both tools, written down below and closed.

## Non-goals (v1)

Charts, conditional formatting, pivot tables, images, data validation, cell comments,
sheet or workbook protection and passwords, Excel tables (ListObject), named-range editing.
Formats other than `.xlsx` / `.xlsm` (`.xls`, `.ods`, `.csv` as input) are also out.
Each of these can be added later as a new `op` without changing the shape.

## Anti-goals (a failure even if it "works")

- **The original file changes.** A new version always gets a new name, which the session picks.
- **Numbers or dates stored as text.** This produces the green-triangle cells that won't sum.
- **Locale-dependent formulas.** `=SOM(A1;A2)` working on one machine and not another. The repo
  was bitten by nl-NL once already (PR #11).
- **Silent macro loss.** Saving an `.xlsm` with VBA as `.xlsx` must not quietly drop the code.
- **Half-applied edits.** An error halfway through the operation list leaves a half-written file.
- **Large responses.** The return value is a small summary, never the sheet content.

## Tool surface

Two new tools, which brings the surface from 38 to **40**. Both are on `ExcelTools`, backed by
`ExcelWorkbookService` and DevExpress `Workbook`, following the pattern in ARCHITECTURE.md.

### `excel_create_workbook(path, sheets[], operations[]?, overwrite=false)`

- `path`: absolute, must end in `.xlsx`. Anything else gives `unsupported_format`; creating an
  `.xlsm` without VBA is pointless.
- `sheets[]`: at least one, written in order.
  - `name`: required, unique, follows Excel's naming rules (≤ 31 chars, none of `: \ / ? * [ ]`).
  - `rows`: array of rows, each an array of cell values (see "Cell values"). Ragged rows are fine.
  - `header`: bool, default `true`. When true, row 1 is the header and gets the **header
    defaults**: bold text, light grey fill (`#D9D9D9`), a thin bottom border, row 1 frozen, and an
    AutoFilter over the written block.
  - `columns[]`: optional per-column settings, each with `column` (letter `"C"` or 1-based number),
    `numberFormat`, `width` (Excel character units), `horizontalAlign` and `wrapText`. These apply
    to the data cells of that column, not the header.
  - Column widths: any column without an explicit `width` is auto-fitted, capped at 60 characters
    so one long text cell doesn't produce a monster column.
- `operations[]`: optional. The same operations as `excel_update_workbook`, applied after the
  sheets are written. This covers needs beyond the defaults, such as a fill colour on a total row
  or merged title cells.
- `overwrite`: default `false`, otherwise an existing file gives `file_exists`. This is the
  existing `PathGuard.RequireWritable`.

### `excel_update_workbook(inputPath, outputPath, operations[], overwrite=false)`

- `inputPath`: existing `.xlsx` / `.xlsm`. It is **never written to**.
- `outputPath`: the new version's name, chosen by the session. It must differ from `inputPath`
  after full-path normalisation, case-insensitive; otherwise `invalid_path`. It must end in
  `.xlsx` or `.xlsm`.
  - `.xlsm` → `.xlsm` keeps the VBA project.
  - An input with a VBA project → `.xlsx` gives `unsupported_format`, with a message saying the
    macros would be lost and to use `.xlsm`.
  - `.xlsx` → `.xlsm` is allowed. It is pointless but harmless.
- `operations[]`: at least one, applied **in order** to the in-memory workbook.
- `overwrite`: applies to `outputPath` only.

### Losing parts on update: save and warn *(decided by the user, 2026-11-04)*

Opening and re-saving a workbook through DevExpress can drop package content it does not carry
over: form-control buttons and their macro assignments (`drawingN.xml` / `vmlDrawingN.vml`),
ActiveX controls, slicers, Power Query / external connections, and some chart types. The VBA
project itself is kept for `.xlsm` → `.xlsm`.

`excel_update_workbook` **saves anyway and reports the loss**:
- Before loading, list the input's package part kinds through the existing
  `Services/Excel/Vba/OpenXmlParts` zip access, plus the form-control/shape macro assignments
  `DrawingMacroExtractor` already finds.
- After saving, do the same on the output.
- Every kind that is present in the input and missing or reduced in the output becomes a
  `warnings[]` entry. Plain-language messages are fine, e.g. `"3 form-control buttons with macro
  assignments on sheet 'Invoer' were not preserved"`.
- A clean update returns `warnings: []`.

The comparison is by part *kind* and count, not byte content, since DevExpress rewrites every part
it keeps. Which kinds to check is left to the builder, starting with the list above.

### Return value (both tools)

```json
{ "outputPath": "C:\\...\\report-v2.xlsx",
  "sheets": [ { "name": "Resultaten", "usedRange": "A1:F120" } ],
  "cellsWritten": 720,
  "operationsApplied": 4,
  "warnings": [] }
```

`warnings` is always present. It is always empty for `excel_create_workbook`.

## Cell values

A cell value is a JSON value, mapped by its JSON type, never by guessing from text:

| JSON | Cell |
|---|---|
| number | numeric cell |
| `true` / `false` | boolean cell |
| `null` | empty: clears the cell in `writeRange` |
| string starting with `=` | formula, English function names, `,` argument separator, `.` decimal (invariant culture) |
| string that is an ISO date `yyyy-MM-dd` | date cell, number format `yyyy-mm-dd` unless the column or a later `formatRange` sets one |
| string that is an ISO datetime `yyyy-MM-dd(T or space)HH:mm[:ss[.fff…]]`, optionally ending in `Z` or `±HH:mm` | date-time cell, format `yyyy-mm-dd hh:mm`, or `yyyy-mm-dd hh:mm:ss` when seconds were given. Fractional seconds are kept in the value. **A `Z` or offset is dropped and the wall-clock time is kept as written**, because Excel has no time zones and PRD's SQL datetimes are local anyway *(designer's call)* |
| any other string | text |
| string starting with `'` | text without the leading apostrophe, for writing text that would otherwise be a formula or date, e.g. `"'=not a formula"` |

The ISO strings are the only text that turns into something else, because JSON has no date type.
A number held in a string (`"12.5"`) **stays text**: the caller decides the type by sending a
JSON number.

The workbook is pinned to `CultureInfo.InvariantCulture`, the same as `LoadWorkbook` today.

**Recalculation** *(designer's call)*:
- **Create:** the whole workbook is recalculated before save, so formula cells carry cached values
  and reading back with `excel_read_sheet` or `excel_export_csv` shows results. Everything in it
  was just written, so nothing can be damaged.
- **Update:** **no** full recalculation by DevExpress. It cannot evaluate VBA user-defined
  functions, external links or some functions, and a full recalc would overwrite Excel's last good
  cached values in cells the caller never touched with `#NAME?` / `#VALUE!`. Instead the saved
  workbook is flagged to recalculate fully when Excel opens it. The trade-off: a formula *written
  by this update* has no cached value until the file is opened in Excel, so mcpOffice's read tools
  show it as empty. This is documented in the tool description. A unit test asserts that an
  untouched formula's cached value survives an update unchanged.

## Operations (closed vocabulary)

Every operation is an object with an `op` string. Sheets are referenced by `sheet` (name).
Ranges and cells are A1 strings (`"B2"`, `"A1:F1"`, `"A:C"` for whole columns, `"3:5"` for whole
rows where it makes sense).

| `op` | Fields | Effect |
|---|---|---|
| `addSheet` | `name`, `position?` (0-based; default end) | new empty sheet |
| `renameSheet` | `sheet`, `newName` | |
| `deleteSheet` | `sheet` | error if it is the last sheet |
| `writeRange` | `sheet`, `startCell`, `rows[][]` | writes values per "Cell values" with no format changes, except dates getting a date format |
| `clearRange` | `sheet`, `range`, `what` = `contents` \| `formats` \| `all` | |
| `formatRange` | `sheet`, `range`, any of: `bold`, `italic`, `underline`, `fontColor`, `fontSize`, `fillColor`, `numberFormat`, `horizontalAlign` (`left`\|`center`\|`right`), `verticalAlign` (`top`\|`middle`\|`bottom`), `wrapText`, `border` (`none`\|`thin`\|`medium`\|`thick`), `borderScope` (`outline`\|`all`, default `all`) | only the given properties change |
| `mergeCells` / `unmergeCells` | `sheet`, `range` | |
| `setColumnWidth` | `sheet`, `columns` (`"A"` or `"A:C"`), `width` **or** `autoFit: true` | |
| `insertRows` / `deleteRows` | `sheet`, `row` (1-based), `count` (default 1) | insert goes *before* `row` |
| `insertColumns` / `deleteColumns` | `sheet`, `column` (letter or number), `count` (default 1) | insert goes *before* `column` |
| `freezePanes` | `sheet`, `cell` (`"A2"` freezes row 1, `"B2"` freezes row 1 and column A), or `null` to unfreeze | |
| `autoFilter` | `sheet`, `range`, or `null` to remove | |

Colours are `#RRGGBB`. Number formats are Excel format strings in invariant form (`"0.00"`,
`"#,##0"`, `"0%"`, `"dd-mm-yyyy"`).

**Atomic:** all operations run against the in-memory workbook and the file is saved only if every
operation succeeds. The first failure throws, naming the operation's index and `op`, for example
`[invalid_operation] operations[3] (formatRange): fillColor 'red' is not #RRGGBB`. Nothing is
written.

**Wire shape:** the operation is a single flat record (`op` plus every optional field) rather than
a polymorphic type hierarchy. The JSON schema the agent sees stays one simple object, and input
binding through MCP SDK 1.2.0 has no polymorphism to get wrong. WORD-003 is the cautionary tale
for output. Validation of which fields an `op` needs is done in code.

## Errors

The existing codes are reused: `file_not_found`, `file_exists`, `invalid_path`,
`unsupported_format`, `sheet_not_found`, `range_too_large`, `parse_error`, `io_error`.

There is **one new code**, `invalid_operation`, for an unknown `op`, missing or contradictory
fields, a malformed range, colour or sheet name, a duplicate sheet name, or deleting the last sheet.

**Limit:** at most **100,000 cells written per call** (sum over `rows` across sheets and
`writeRange`s), otherwise `range_too_large`. This is far above what fits in an agent's message;
it stops a runaway call rather than limiting real use.

## Testing and acceptance

- **Unit tests** (`tests/mcpOffice.Tests/Excel/`) against the service, **reading the result back
  through the existing read path** (`ReadSheet`, `ListSheets`). This proves what an agent would
  actually see. Cover:
  - every row of the cell-value table, including that `"12.5"` stays text and `'` escapes
  - every `op`
  - the header defaults
  - atomicity: a failing op 3 leaves no output file
  - `inputPath` unchanged byte-for-byte after an update
  - the VBA-preserving `.xlsm` round trip, reusing `tests/fixtures/sample-with-macros.xlsm`, and
    the `.xlsm` → `.xlsx` refusal
  - lost-part detection: an input with a form-control button, either `synthetic-vba.xlsm` or a
    corpus file, gives a warning if DevExpress drops it, and a clean update gives `warnings: []`
  - an untouched formula's cached value surviving an update
  - an nl-NL thread culture producing identical formulas and numbers
- **Integration test:** one create and one update **through the stdio transport**. The `rows` and
  `operations` input binding is exactly what unit tests cannot see. `ToolSurfaceTests` goes to 40.
- **Live acceptance:** the live `office` server in a session. Create a workbook from a real PRD
  query result, then make a new version of a corpus workbook in `C:\Projects\mcpOffice-samples`,
  and open both in Excel.
- **Docs:** README tool list (38 → 40), `docs/usage.md` example, ARCHITECTURE.md unchanged apart
  from the count if it is mentioned.

## Left to the builder

- How the service is split internally: an operation applier class versus methods on the service.
  Pure parts (A1 parsing, value mapping, colour parsing) get string-level unit tests, like the VBA
  classes.
- The auto-fit approach (DevExpress `AutoFitColumns` versus measuring text length) as long as the
  60-character cap holds.
- Exact wording of tool descriptions, which should be written for an LLM caller per ARCHITECTURE.md.
- Whether `writeRange` grows the AutoFilter range of a sheet created with `header: true`.
  Sensible: no, filters are explicit after creation.
