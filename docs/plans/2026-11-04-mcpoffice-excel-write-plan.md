# Excel write tools (create / new version) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `excel_create_workbook` and `excel_update_workbook` so a session can hand over a formatted `.xlsx` in one call, or a new version of an existing workbook under a new name.

**Architecture:** Two thin `[McpServerTool]` methods on `ExcelTools` delegate to two new `ExcelWorkbookService` methods. Those build the workbook in memory with DevExpress `Workbook`, apply a flat list of operations, and save atomically. The pure pieces are A1 parsing, JSON-to-cell value mapping and package-part comparison. They live under `Services/Excel/Write/` with string-level unit tests. DevExpress work stays in a writer and an operation applier.

**Tech Stack:** .NET 10, C#, DevExpress Spreadsheet (`DevExpress.Document.Processor` 26.1.6), ModelContextProtocol 1.2.0, `System.Text.Json` (`JsonElement` for heterogeneous cell values), xUnit `Assert.*`.

**Spec:** `docs/plans/2026-11-04-mcpoffice-excel-write-design.md`. Read it first; this plan argues from it.

## Global Constraints

- Every tool error is an `McpException` whose message starts with `[code]`. Codes come from `ErrorCode.cs`; the only new one is `invalid_operation`.
- The workbook is always pinned to `workbook.Options.Culture = CultureInfo.InvariantCulture`, before load or before writing.
- Create outputs `.xlsx` only. Update reads `.xlsx` / `.xlsm` and writes `.xlsx` / `.xlsm`.
- `outputPath` must differ from `inputPath` after `Path.GetFullPath`, compared case-insensitively.
- `inputPath` is never opened for writing.
- Default for `overwrite` is `false`, enforced with `PathGuard.RequireWritable`.
- At most **100,000** cells written per call, counted across `sheets[].rows` and every `writeRange.rows`, before anything is applied.
- Auto-fit column width is capped at **60** characters.
- Header defaults: bold, fill `#D9D9D9`, thin bottom border, row 1 frozen, AutoFilter over the written block.
- Create runs `workbook.CalculateFull()` before save. Update does **not** recalculate. It sets `workbook.DocumentSettings.Calculation.FullCalculationOnLoad = true`.
- Saving is atomic: write to a temp file in the output's directory, then `File.Move(temp, output, overwrite)`. On any exception, delete the temp file and leave the output absent.
- Both tools always return `warnings` (an empty list for create).
- File-scoped namespaces, nullable enabled. Tools set `[McpServerTool(Name = "...")]` explicitly. Nothing is written to stdout.
- Tests use xUnit `Assert.*` only; there is no FluentAssertions. Fixtures are generated programmatically, except the existing `tests/fixtures/*.xlsm`.
- Conventional Commits, on branch `feat/card-2853-excel-create`.

## Review Focus

Inputs and failure modes the spec implies but doesn't spell out. Each line names the test that pins it and the task that owns it.

1. **A Dutch-style formula** `"=SOM(A1;A2)"` or `"=SUM(A1;A2)"` must give `invalid_operation` naming the cell, not a silent `#NAME?`. Test: `Formula_with_semicolon_separator_is_invalid_operation`, Task 3.
2. **A JSON object or array as a cell value** (`{"a":1}`, `[1,2]`) must give `invalid_operation`, never be written as `"System.Text.Json..."` text. Test: `Object_or_array_value_is_rejected`, Task 2.
3. **Sheet names that differ only by case** (`"Data"` and `"data"`) count as duplicates, the same as in Excel, and give `invalid_operation`. Test: `Duplicate_sheet_name_case_insensitive_is_invalid_operation`, Task 3.
4. **Text longer than 32,767 characters** (Excel's cell limit) gives `invalid_operation` and is not truncated. Test: `Text_over_32767_chars_is_rejected`, Task 2.
5. **An `outputPath` in a folder that doesn't exist yet** gets the folder created, and a missing `inputPath` gives `file_not_found` with no output. Test: `Update_creates_output_directory_and_missing_input_is_file_not_found`, Task 4.

---

### Task 1: Write vocabulary, parsers and the new error code

**Files:**
- Create: `src/mcpOffice/Models/ExcelWriteSpecs.cs`
- Create: `src/mcpOffice/Models/ExcelWriteResult.cs`
- Create: `src/mcpOffice/Services/Excel/Write/ExcelWriteParsers.cs`
- Create: `src/mcpOffice/Services/Excel/Write/WriteValidationException.cs`
- Modify: `src/mcpOffice/ErrorCode.cs`, `src/mcpOffice/ToolError.cs`
- Test: `tests/mcpOffice.Tests/Excel/Write/ExcelWriteParsersTests.cs`

**Interfaces:**
- Produces (Models, `namespace McpOffice.Models`):
  - `sealed record ExcelSheetSpec(string Name, JsonElement[][]? Rows = null, bool Header = true, ExcelColumnSpec[]? Columns = null)`
  - `sealed record ExcelColumnSpec(string Column, string? NumberFormat = null, double? Width = null, string? HorizontalAlign = null, bool? WrapText = null)`
  - `sealed record ExcelOperation(string Op, string? Sheet = null, string? Name = null, int? Position = null, string? NewName = null, string? StartCell = null, JsonElement[][]? Rows = null, string? Range = null, string? What = null, bool? Bold = null, bool? Italic = null, bool? Underline = null, string? FontColor = null, double? FontSize = null, string? FillColor = null, string? NumberFormat = null, string? HorizontalAlign = null, string? VerticalAlign = null, bool? WrapText = null, string? Border = null, string? BorderScope = null, string? Columns = null, double? Width = null, bool? AutoFit = null, int? Row = null, string? Column = null, int? Count = null, string? Cell = null)`. This is a flat record; the spec's "Wire shape" section explains why.
  - `sealed record ExcelWrittenSheet(string Name, string UsedRange)`
  - `sealed record ExcelWriteResult(string OutputPath, IReadOnlyList<ExcelWrittenSheet> Sheets, int CellsWritten, int OperationsApplied, IReadOnlyList<string> Warnings)`
- Produces (`internal static class ExcelWriteParsers`, `namespace McpOffice.Services.Excel.Write`). Every method throws `WriteValidationException(string message)` on bad input:
  - `int ParseColumn(string letterOrNumber)`: `"A"`→1, `"c"`→3, `"AA"`→27, `"3"`→3. Range is 1..16384.
  - `string ColumnLetter(int column)`: 1-based.
  - `(int Row, int Column) ParseCell(string a1)`: 1-based. `"B2"`→(2,2). Rows 1..1048576.
  - `void ValidateSheetName(string name)`: not blank, ≤ 31 chars, none of `: \ / ? * [ ]`, doesn't start or end with `'`.
  - `System.Drawing.Color ParseHexColor(string hex)`: only `#RRGGBB`, case-insensitive.
- Produces (`ToolError`):
  - `InvalidOperation(string where, string detail)` → `[invalid_operation] {where}: {detail}`, where `where` is like `operations[3] (formatRange)` or `sheets[1]`.
  - `UnsupportedWorkbookFormat(string detail)` → `[unsupported_format] {detail}`.
  - `TooManyCells(int count, int max)` → `[range_too_large] This call writes {count} cells, which exceeds the limit of {max}.`

- [ ] **Step 1: Write the failing tests.** Add `[Theory]` cases for each parser: the values listed above, plus rejections of `"0"`, `"XFE"`, `"A0"`, `"1A"`, `""`, a 32-char name, `"a/b"`, `"'x"`, `"red"`, `"#FFF"` and `"FF0000"`. Each rejection asserts `Assert.Throws<WriteValidationException>`. Also add `ToolError_InvalidOperation_prefix`, asserting the message starts with `[invalid_operation] operations[3] (formatRange): `.
- [ ] **Step 2: Run** `dotnet test tests/mcpOffice.Tests --filter "FullyQualifiedName~Excel.Write"`. Expected: build FAIL because the types are missing.
- [ ] **Step 3: Implement** the records, the parsers, `WriteValidationException : Exception`, `ErrorCode.InvalidOperation = "invalid_operation"` and the three `ToolError` helpers.
- [ ] **Step 4: Run** the same filter. Expected: PASS.
- [ ] **Step 5: Commit** `feat(excel): write-tools vocabulary, A1 parsers, invalid_operation code`.

### Task 2: JSON → cell value mapping

**Files:**
- Create: `src/mcpOffice/Services/Excel/Write/CellValueMapper.cs`
- Test: `tests/mcpOffice.Tests/Excel/Write/CellValueMapperTests.cs`

**Interfaces:**
- Produces: `internal enum MappedCellKind { Empty, Number, Boolean, Text, Formula, DateTime }`, `internal sealed record MappedCell(MappedCellKind Kind, object? Value, string? DefaultNumberFormat)` and `internal static class CellValueMapper { public static MappedCell Map(JsonElement value); }`. It throws `WriteValidationException`.

- [ ] **Step 1: Write the failing tests.** Parse each input with `JsonDocument.Parse(...).RootElement`, then assert:
  - `12.5` → (Number, 12.5d, null)
  - `true` → (Boolean, true, null)
  - `null` → (Empty, null, null)
  - `"=SUM(A1:A2)"` → (Formula, `"=SUM(A1:A2)"`, null)
  - `"2026-11-04"` → (DateTime, `new DateTime(2026,11,4)`, `"yyyy-mm-dd"`)
  - `"2026-11-04T10:15"` and `"2026-11-04 10:15"` → DateTime 10:15, `"yyyy-mm-dd hh:mm"`
  - `"2026-11-04T10:15:30.123"` → DateTime with 123 ms, `"yyyy-mm-dd hh:mm:ss"`
  - `"2026-11-04T10:15:30Z"` and `"2026-11-04T10:15:30+01:00"` → DateTime **10:15:30** (offset dropped), `"yyyy-mm-dd hh:mm:ss"`
  - `"12.5"` → (Text, `"12.5"`, null)
  - `"'=not a formula"` → (Text, `"=not a formula"`)
  - `"'2026-11-04"` → (Text, `"2026-11-04"`)
  - `"2026-13-45"` → (Text, unchanged)
  - `"hello"` → Text
  - **Review Focus 2:** `Object_or_array_value_is_rejected`: `{"a":1}` and `[1,2]` throw `WriteValidationException`.
  - **Review Focus 4:** `Text_over_32767_chars_is_rejected`: a 32,768-char string throws, and 32,767 maps to Text.
- [ ] **Step 2: Run** `--filter "FullyQualifiedName~CellValueMapper"`. Expected: build FAIL.
- [ ] **Step 3: Implement `CellValueMapper.Map`.**
  - Parse dates with `DateTime.TryParseExact` using `CultureInfo.InvariantCulture` and an explicit format list. Remove any trailing `Z` or `±HH:mm` with a regex before parsing.
  - Never use `DateTime.Parse`, because it is culture- and heuristics-dependent.
  - The `'` escape is checked first, then `=`, then the date patterns.
- [ ] **Step 4: Run.** Expected: PASS.
- [ ] **Step 5: Commit** `feat(excel): JSON-to-cell value mapping for write tools`.

### Task 3: `CreateWorkbook`: sheets, header defaults, columns, recalc, atomic save

**Files:**
- Create: `src/mcpOffice/Services/Excel/Write/WorkbookWriter.cs`
- Modify: `src/mcpOffice/Services/Excel/IExcelWorkbookService.cs`, `src/mcpOffice/Services/Excel/ExcelWorkbookService.cs`
- Test: `tests/mcpOffice.Tests/Excel/Write/CreateWorkbookTests.cs`

**Interfaces:**
- Consumes: Task 1 records and parsers, Task 2 `CellValueMapper`.
- Produces:
  - `internal static class WorkbookWriter`:
    - `int WriteRows(Worksheet ws, int startRow, int startColumn, JsonElement[][] rows)`: 1-based start, returns cells written, and applies the `DefaultNumberFormat` from each `MappedCell`. A formula DevExpress rejects throws `WriteValidationException("cell B2: <reason>")`.
    - `void WriteSheet(Worksheet ws, ExcelSheetSpec spec)`: rows, then header defaults if `Header`, then `Columns[]`, then auto-fit with the 60 cap for columns without a `Width`.
    - `void SaveAtomic(Workbook wb, string outputPath)`: chooses `DocumentFormat.Xlsx` or `DocumentFormat.Xlsm` from the extension.
    - `ExcelWriteResult Summarize(Workbook wb, string outputPath, int cellsWritten, int operationsApplied, IReadOnlyList<string> warnings)`
  - `IExcelWorkbookService.CreateWorkbook(string path, IReadOnlyList<ExcelSheetSpec> sheets, IReadOnlyList<ExcelOperation>? operations, bool overwrite) → ExcelWriteResult`. In this task, a non-empty `operations` throws `invalid_operation` "operations not supported yet"; Task 4 wires it up.

- [ ] **Step 1: Write the failing tests.** Each test creates a workbook in a temp directory, then checks it through `new ExcelWorkbookService().ReadSheet(...)` and, for formatting, through a DevExpress `Workbook` loaded in the test.
  - `Create_writes_typed_values_readable_back`: rows `[["Name","Amount","Date","Total"],["Ada",40,"2026-11-04","=B2*2"]]`. Read-back has B2 `ValueType == "number"`, C2 `"datetime"`, D2 `"number"` with value `80d` (create recalculates), and A2 `"text"`.
  - `Create_applies_header_defaults`: A1 bold, fill `#D9D9D9`, bottom border thin, `ws.ActiveView.FrozenCell`/freeze state showing row 1 frozen, AutoFilter range `A1:D2`.
  - `Create_without_header_has_no_bold_freeze_or_filter`.
  - `Create_autofit_caps_at_60_chars`: a 200-char text column has width ≤ 60 characters (convert with the DevExpress column width unit), and a short column is narrower than 60.
  - `Create_applies_column_specs_to_data_cells_not_header`: `{"column":"B","numberFormat":"0.00","width":12}`. B2's format is `0.00`, B1's isn't, and B's width is 12.
  - `Create_multiple_sheets_in_order` and `Create_sheet_with_empty_rows_creates_empty_sheet`.
  - `Create_rejects_non_xlsx_path`: `.xlsm` and `.csv` give `[unsupported_format]`.
  - `Create_existing_file_without_overwrite_is_file_exists`.
  - **Review Focus 3:** `Duplicate_sheet_name_case_insensitive_is_invalid_operation`: `Data` + `data` gives `[invalid_operation] sheets[1]`.
  - `Invalid_sheet_name_is_invalid_operation`.
  - **Review Focus 1:** `Formula_with_semicolon_separator_is_invalid_operation`: `"=SUM(A1;A2)"` gives `[invalid_operation]` with `B2` in the message, or the cell where it was placed, and no file at `path`.
  - `Over_100000_cells_is_range_too_large_and_writes_nothing`.
  - `Create_is_culture_independent`: run the first test with `CultureInfo.CurrentCulture = new CultureInfo("nl-NL")` (restore it in `finally`). The values and formula are identical.
- [ ] **Step 2: Run** `--filter "FullyQualifiedName~CreateWorkbook"`. Expected: build FAIL.
- [ ] **Step 3: Implement** `WorkbookWriter` and `ExcelWorkbookService.CreateWorkbook`.
  - Order: guard paths, count cells and check the limit, validate sheet names (case-insensitive set), then `new Workbook()` with invariant culture.
  - Rename the default sheet to `sheets[0].Name` and add the rest.
  - Then `WriteSheet` for each, `CalculateFull()`, `SaveAtomic`, `Summarize`.
  - Wrap with `catch (WriteValidationException ex)` → `ToolError.InvalidOperation("sheets[i]", ex.Message)`, then the usual `catch (Exception ex) when (ex is not McpException)` → `ToolError.IoError`/`ParseError`.
  - Use DevExpress calls `ws.FreezeRows(0)`, `ws.AutoFilter.Apply(range)` and `ws.Columns.AutoFit(first, last)`, then clamp each width.
- [ ] **Step 4: Run.** Expected: PASS. Then `dotnet build`: 0 warnings.
- [ ] **Step 5: Commit** `feat(excel): CreateWorkbook with typed values and header defaults`.

### Task 4: `UpdateWorkbook` and the operation applier: sheet ops, writeRange, clearRange

**Files:**
- Create: `src/mcpOffice/Services/Excel/Write/OperationApplier.cs`
- Modify: `IExcelWorkbookService.cs`, `ExcelWorkbookService.cs` (`CreateWorkbook` now applies `operations`)
- Test: `tests/mcpOffice.Tests/Excel/Write/UpdateWorkbookTests.cs`

**Interfaces:**
- Consumes: `WorkbookWriter.WriteRows`, `SaveAtomic`, `Summarize`.
- Produces:
  - `internal static class OperationApplier { public static int Apply(Workbook wb, IReadOnlyList<ExcelOperation> ops); }` returns cells written. A failure at index `i` throws `ToolError.InvalidOperation($"operations[{i}] ({op.Op})", detail)`. A missing sheet throws `ToolError.SheetNotFound` with the message prefixed `operations[i] (op): `. An unknown `op` lists the valid ones.
  - Later tasks add cases to the same `switch` on `Op`, which is **case-sensitive** and matches the spec's spelling.
  - `IExcelWorkbookService.UpdateWorkbook(string inputPath, string outputPath, IReadOnlyList<ExcelOperation> operations, bool overwrite) → ExcelWriteResult`. `Warnings` is `[]` until Task 7.
- Ops in this task: `addSheet`, `renameSheet`, `deleteSheet`, `writeRange`, `clearRange`.

- [ ] **Step 1: Write the failing tests.** The input is created with `TestExcelWorkbooks.Create` and includes a formula with a cached value.
  - `Update_writes_new_version_and_leaves_input_byte_identical`: SHA-256 of the input before and after are equal, and the output contains the written values.
  - `Update_output_equal_to_input_is_invalid_path`, including the same path in different case.
  - `Update_does_not_recalculate_untouched_formulas`: an input formula's cached value survives unchanged, and the output's `DocumentSettings.Calculation.FullCalculationOnLoad` is `true`.
  - `Failing_operation_writes_nothing`: ops `[writeRange ok, writeRange ok, formatRange-unknown-sheet]`. Expect `sheet_not_found` with `operations[2]` in the message, and no output file.
  - `Unknown_op_is_invalid_operation_listing_valid_ops`.
  - Sheet ops: `addSheet` with and without `position`, `renameSheet`, `deleteSheet`. `deleteSheet` on the last sheet gives `invalid_operation`, and a duplicate `addSheet` name gives `invalid_operation`.
  - `writeRange_null_clears_cell`. `clearRange` with `what` = `contents`, `formats`, `all`; a bad `what` gives `invalid_operation`.
  - `Create_applies_operations_after_sheets`.
  - **Review Focus 5:** `Update_creates_output_directory_and_missing_input_is_file_not_found`.
- [ ] **Step 2: Run** `--filter "FullyQualifiedName~UpdateWorkbook"`. Expected: build FAIL.
- [ ] **Step 3: Implement.**
  - `UpdateWorkbook`, in order:
    1. `PathGuard.RequireExists(inputPath)`
    2. extension checks
    3. the full-path inequality check
    4. `RequireWritable(outputPath, overwrite)`
    5. the cell-limit count
    6. load with invariant culture
    7. `OperationApplier.Apply`
    8. `FullCalculationOnLoad = true`
    9. `SaveAtomic`
  - **Leave the `.xlsm` / VBA rules for Task 7.** Here, any `.xlsm` input or output that is not paired `.xlsm` → `.xlsm` is passed through unchanged.
- [ ] **Step 4: Run.** Expected: PASS.
- [ ] **Step 5: Commit** `feat(excel): UpdateWorkbook with sheet, writeRange and clearRange operations`.

### Task 5: Formatting operations

**Files:**
- Modify: `src/mcpOffice/Services/Excel/Write/OperationApplier.cs`
- Test: `tests/mcpOffice.Tests/Excel/Write/FormatOperationsTests.cs`

**Interfaces:**
- Consumes: `OperationApplier.Apply` switch, `ExcelWriteParsers.ParseHexColor`, `ParseColumn`.
- Ops: `formatRange`, `mergeCells`, `unmergeCells`, `setColumnWidth`, `freezePanes`, `autoFilter`.

- [ ] **Step 1: Write the failing tests.** Run each through `UpdateWorkbook`, then inspect with DevExpress.
  - `formatRange_sets_only_given_properties`: set `fillColor` on a range that was already bold. It stays bold and the fill is set.
  - One test per property: font (bold, italic, underline, colour, size), fill, numberFormat, horizontal and vertical align, wrapText.
  - `border` with `borderScope` `all` versus `outline`: for `outline` on `A1:B2`, A1's top and left edges are set and its inner right edge isn't.
  - Bad enum values (`horizontalAlign:"middle"`, `border:"dotted"`) and `fillColor:"red"` give `invalid_operation` with the op index.
  - `mergeCells`/`unmergeCells` round trip.
  - `setColumnWidth`: `"A:C"` width 20 sets all three; `autoFit:true` respects the 60 cap; giving **both** `width` and `autoFit`, or neither, gives `invalid_operation`.
  - `freezePanes`: `"A2"` freezes rows only, `"B2"` freezes a row and a column, `null` unfreezes.
  - `autoFilter`: a range applies it, `null` disables it.
- [ ] **Step 2: Run** `--filter "FullyQualifiedName~FormatOperations"`. Expected: FAIL because the ops are unknown.
- [ ] **Step 3: Implement** the six cases. `freezePanes` maps cell (r,c) to `ws.FreezePanes(r-2, c-2)` / `FreezeRows` / `FreezeColumns` as the zero-based counts require, or `ws.UnfreezePanes()` for null.
- [ ] **Step 4: Run.** Expected: PASS.
- [ ] **Step 5: Commit** `feat(excel): formatting operations for write tools`.

### Task 6: Row and column insert/delete

**Files:**
- Modify: `OperationApplier.cs`
- Test: `tests/mcpOffice.Tests/Excel/Write/RowColumnOperationsTests.cs`

**Interfaces:** the ops `insertRows`, `deleteRows`, `insertColumns`, `deleteColumns`, with fields `row` / `column` and `count` (default 1).

- [ ] **Step 1: Write the failing tests.**
  - `insertRows row:2 count:2` moves the old A2 to A4.
  - `deleteRows` moves the values up.
  - A formula `=SUM(B2:B3)` that refers to a shifted range is adjusted by DevExpress; assert the formula text after `insertRows` above it.
  - The column equivalents, with `column` given as `"B"` and as `"2"`.
  - `count:0`, a negative count, or a missing `row` give `invalid_operation`.
- [ ] **Step 2: Run.** Expected: FAIL.
- [ ] **Step 3: Implement** with `ws.Rows.Insert(index, count)` / `ws.Rows.Remove(index, count)` and the `Columns` equivalents. Indexes are zero-based.
- [ ] **Step 4: Run.** Expected: PASS.
- [ ] **Step 5: Commit** `feat(excel): row and column insert/delete operations`.

### Task 7: `.xlsm` rules and lost-part warnings

**Files:**
- Create: `src/mcpOffice/Services/Excel/Write/PackagePartInventory.cs`
- Modify: `ExcelWorkbookService.UpdateWorkbook`
- Test: `tests/mcpOffice.Tests/Excel/Write/PackagePartInventoryTests.cs`, `tests/mcpOffice.Tests/Excel/Write/UpdateWorkbookMacroTests.cs`

**Interfaces:**
- Consumes: `OpenXmlParts.ListSheets`, `OpenXmlParts.ReadEntryText`, `DrawingMacroExtractor.FromDrawingXml`/`FromVmlDrawing`, and `TestOpenXmlPackages.Build`/`BuildWorkbookWithDrawings` for pure tests.
- Produces:
  - `internal static class PackagePartInventory`
  - `IReadOnlyDictionary<string,int> Take(ZipArchive zip)`. Its kinds are a closed list:
    - `vbaProject` (has `xl/vbaProject.bin`)
    - `macroButtons:<sheet>` (shape and form-control macro assignments per sheet)
    - `activeX` (`xl/activeX/*`)
    - `slicers` (`xl/slicers/*`)
    - `connections` (`xl/connections.xml`)
    - `queryTables` (`xl/queryTables/*`)
    - `charts` (`xl/charts/chart*.xml`)
    - `pivotTables` (`xl/pivotTables/*`)
  - `IReadOnlyList<string> Compare(IReadOnlyDictionary<string,int> before, IReadOnlyDictionary<string,int> after)`: one plain-language warning per kind whose count dropped, e.g. `"3 form-control buttons with macro assignments on sheet 'Invoer' were not preserved"`.

- [ ] **Step 1: Write the failing tests.**
  - Pure: `Take` counts each kind from a `TestOpenXmlPackages`-built zip. `Compare` with equal maps gives `[]`. Dropped buttons give the exact sentence above.
  - Service:
    - `Xlsm_to_xlsm_keeps_vba_project`: copy `tests/fixtures/sample-with-macros.xlsm` to temp and update to `.xlsm`. `new ExcelWorkbookService().ExtractVba(output)` returns the same module names.
    - `Xlsm_with_vba_to_xlsx_is_unsupported_format`: the message mentions macros and `.xlsm`, and no output exists.
    - `Xlsx_to_xlsm_is_allowed`.
    - `Clean_update_returns_empty_warnings`.
    - `Update_reports_lost_macro_buttons_if_devexpress_drops_them`: use `tests/fixtures/synthetic-vba.xlsm`. If DevExpress *keeps* its buttons, assert `warnings` is empty and the inventory counts are equal; if it drops them, assert the warning. Write it as an equality between `Compare(Take(in), Take(out))` and the returned `Warnings`, so it holds either way.
- [ ] **Step 2: Run** `--filter "FullyQualifiedName~PackagePartInventory|FullyQualifiedName~UpdateWorkbookMacro"`. Expected: FAIL.
- [ ] **Step 3: Implement.**
  - In `UpdateWorkbook`, before load: if the input has `vbaProject` and the output is `.xlsx`, throw `ToolError.UnsupportedWorkbookFormat(...)`.
  - Take the inventory of the input, and of the output after `SaveAtomic`. Both are read-only `ZipFile.OpenRead`.
  - Put `Compare` into `Warnings`.
- [ ] **Step 4: Run.** Expected: PASS.
- [ ] **Step 5: Commit** `feat(excel): keep VBA on .xlsm updates and warn about lost package parts`.

### Task 8: MCP tools, tool surface, transport test

**Files:**
- Modify: `src/mcpOffice/Tools/ExcelTools.cs`, `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`
- Create: `tests/mcpOffice.Tests.Integration/ExcelWriteWorkflowTests.cs`

**Interfaces:**
- Produces:
  - `[McpServerTool(Name = "excel_create_workbook")] public static object ExcelCreateWorkbook(string path, ExcelSheetSpec[] sheets, ExcelOperation[]? operations = null, bool overwrite = false)`
  - `[McpServerTool(Name = "excel_update_workbook")] public static object ExcelUpdateWorkbook(string inputPath, string outputPath, ExcelOperation[] operations, bool overwrite = false)`
  - Both are one-line delegations. The `[Description]`s are written for an LLM. Together they must state:
    - the cell-value rules (JSON type decides; ISO dates; `=` formulas in English with `,`; `'` escape)
    - the op list
    - that the original is never modified
    - that update doesn't recalculate, so newly written formulas show empty in mcpOffice reads until opened in Excel
    - `warnings[]`

- [ ] **Step 1: Write the failing tests.**
  - `ToolSurfaceTests`: the expected catalog gains both names, and the count goes from 38 to 40.
  - `Create_then_update_via_stdio`, through `ServerHarness` and `CallToolAsync`:
    - create a two-sheet workbook with mixed JSON types (number, bool, null, ISO date, formula)
    - call `excel_update_workbook` with `writeRange` + `formatRange` to a new path
    - call `excel_read_sheet` on the output
    - assert the types and values and that `warnings` is `[]`
    - also assert the input's hash is unchanged
  - This is the test that proves `JsonElement[][]` and the flat `ExcelOperation` bind through SDK 1.2.0.
- [ ] **Step 2: Run** `dotnet build` then `dotnet test tests/mcpOffice.Tests.Integration`. Expected: FAIL because the tools are missing.
- [ ] **Step 3: Implement** the two tool methods.
- [ ] **Step 4: Run** the full `dotnet test`. Expected: all pass, and the build has 0 warnings.
- [ ] **Step 5: Commit** `feat(excel): excel_create_workbook and excel_update_workbook tools`.

### Task 9: Docs and live acceptance

**Files:**
- Modify: `README.md` (status line and tool count 38 → 40; Excel list gains both tools with one-line summaries; roadmap item 8 "Excel write" ✓)
- Modify: `docs/usage.md` (one create example and one update example in JSON)
- Modify: `SESSION_HANDOFF.md` (this session)

- [ ] **Step 1:** Update the docs. Grep the repo for `38 tools` and `38` tool-count mentions to catch them all.
- [ ] **Step 2: Live acceptance.**
  1. Rebuild using the kill-and-build command in SESSION_HANDOFF, then `/mcp`.
  2. In the session, use the live `office` server to create a workbook from a real PRD query result: run a few dozen rows from `mcp__mcpsql__execute_query` against `PrdConnection`, with dates and numbers, through `excel_create_workbook`.
  3. Run `excel_update_workbook` on a copy-sourced corpus workbook in `C:\Projects\mcpOffice-samples`, preferably an `.xlsm` with buttons, to a new name.
  4. Open both in Excel. Check that numbers aren't flagged as text, dates are dates, the header is frozen and filtered, and the macros run. Record any `warnings` returned.
  5. Card any gap found rather than working around it.
- [ ] **Step 3: Commit** `docs: Excel write tools in README/usage, handoff`, push the branch, open a PR, squash-merge after review.
- [ ] **Step 4:** Move card #2853 to Review with a conclusion citing the merge SHA. The share redeploy is batched with this, per the user. It happens after merge and is tracked on #2854.
