# Excel POC Handoff — 2026-05-01

## Branch

`poc/excel-tools`

## Goal

Build an Excel POC with two tracks:

1. Read workbook/sheet data into JSON.
2. Extract workbook structure and VBA/macro source for future Excel-to-C# conversion workflows.

## Sample File

Use this macro workbook for the second track:

```text
C:\temp\macro\Air - Labware.xlsm
```

Confirmed locally:

- file exists;
- contains `xl/vbaProject.bin`;
- has 69 worksheets;
- contains `xl/calcChain.xml`;
- largest sampled sheet is `32+32+32 MPN tabel` with dimension `A1:M35939`.

## Current Findings

- DevExpress Spreadsheet Document API supports loading XLSM, but official docs say macros cannot be executed or modified.
- Excel COM can open the sample read-only with macros disabled and sees 69 worksheets, but did not expose VBA components on this machine.
- Static VBA extraction should not rely on Excel COM. Preferred path is ZIP + OLE structured storage + VBA decompression.
- Python/oletools is not available locally.

## Next Implementation Steps

1. ✅ Add `DevExpress.Document.Processor` to server and unit test projects.
2. ✅ Create Excel DTOs and `IExcelWorkbookService`.
3. ✅ Implement `excel_list_sheets(path)`.
4. ✅ Add unit tests using generated `.xlsx` fixtures.
5. ✅ Add MCP tool wrapper and integration test.
6. Implement `excel_read_sheet(path, sheetName?, sheetIndex?, range?, includeFormulas=true, includeFormats=false, maxCells=50000)`.
7. Spike static `excel_extract_vba(path)` against `C:\temp\macro\Air - Labware.xlsm`.

## Current Verification

`dotnet test --nologo` passes: 42/42 tests (35 unit + 7 integration).

## Design Doc

See:

```text
docs/plans/2026-05-01-mcpoffice-excel-poc-design.md
```
