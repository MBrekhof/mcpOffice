# mcpOffice Excel POC Design

**Date:** 2026-05-01
**Branch:** `poc/excel-tools`
**Status:** Draft

## Goal

Add an Excel (`.xlsx` / `.xlsm`) proof of concept to mcpOffice with two deliberately different use cases:

1. **Data extraction:** read an Excel workbook and return all data from a sheet, or a bounded range, in an agent-friendly JSON shape.
2. **Conversion analysis:** retrieve workbook structure, formulas, defined names, and VBA/macro source code so an agent can help convert an Excel/VBA solution into a maintainable C# program.

The second use case is not just "read cell values"; it needs to expose intent: sheets, dependencies, formula hotspots, event handlers, modules, procedures, and places where VBA talks to Excel objects.

## Sample Workbook

Macro/structure spike file:

```text
C:\temp\macro\Air - Labware.xlsm
```

Initial static inspection:

- File exists locally, size ~2.8 MB.
- Open XML package contains `xl/vbaProject.bin` (~1.1 MB).
- Workbook has 69 worksheets.
- Workbook contains `xl/calcChain.xml` (~1.0 MB).
- Largest sampled sheet: `32+32+32 MPN tabel`, dimension `A1:M35939`, ~251k populated cells, ~36k formulas.
- Excel COM can open the workbook read-only with automation security forced off and sees 69 worksheets, but did not expose VBA components on this machine. Do not make COM/Trust Center access the primary macro extraction path.

## Technical Findings

### DevExpress Spreadsheet API

DevExpress Spreadsheet Document API can load Excel workbooks through `DevExpress.Spreadsheet.Workbook.LoadDocument`. Official docs list XLSM as supported with limited macro support: macros cannot be executed or modified.

This is enough for workbook/sheet/cell/formula extraction, but not enough by itself for VBA source extraction.

Expected package:

```xml
<PackageReference Include="DevExpress.Spreadsheet.Core" Version="25.2.5" />
```

### VBA Extraction

For `.xlsm`, macro code lives in `xl/vbaProject.bin`, which is an OLE compound file embedded in the ZIP package. The VBA source streams are compressed. A safe extractor should:

- open `.xlsm` as ZIP;
- locate `xl/vbaProject.bin`;
- parse OLE structured storage;
- locate VBA module streams;
- decompress VBA source text;
- return modules/classes/forms without opening Excel or executing macros.

Candidate implementation paths:

- **Preferred POC path:** use `OpenMcdf` to read OLE storage and implement the VBA dir/module/decompression logic in C#.
- **Fallback path:** shell out to a local `olevba`/`oletools` executable if Python tooling is installed and explicitly configured. Not preferred for the server because Python is not present on this machine.
- **Avoid as primary path:** Excel COM automation. It depends on Excel being installed, Trust Center settings, and safe automation configuration.

## Initial Tool Surface

### Track A — Data Extraction

#### `excel_list_sheets(path)`

Returns workbook sheets in order.

```json
[
  {
    "index": 0,
    "name": "WO",
    "visible": true,
    "kind": "worksheet",
    "usedRange": "A1:Z100",
    "rowCount": 100,
    "columnCount": 26
  }
]
```

#### `excel_read_sheet(path, sheetName?, sheetIndex?, range?, includeFormulas=true, includeFormats=false, maxCells=50000)`

Returns cell data for a worksheet or range. If `range` is omitted, use the used range. `maxCells` prevents accidentally streaming giant sheets without intent.

```json
{
  "sheet": "WO",
  "range": "A1:D3",
  "truncated": false,
  "rows": [
    ["Sample", "Value", "Unit", "Formula"],
    ["A", 42, "mg/L", null]
  ],
  "cells": [
    {
      "address": "B2",
      "value": 42,
      "valueType": "number",
      "formula": null,
      "displayText": "42",
      "numberFormat": "0"
    }
  ]
}
```

Default response should include both:

- `rows` for easy agent consumption;
- `cells` for addresses, formulas, and metadata.

#### `excel_get_metadata(path)`

Returns author/title/created/modified plus workbook counts.

### Track B — Structure + Macro Analysis

#### `excel_get_structure(path, includeSheets=true, includeFormulas=true, includeDefinedNames=true)`

Returns workbook-level structure:

- sheets and visibility;
- used ranges;
- formula counts by sheet;
- defined names;
- external connections if discoverable;
- table/pivot/chart counts if available through DevExpress or Open XML.

#### `excel_list_formulas(path, sheetName?, includeValues=false, maxFormulas=10000)`

Returns formula cells with sheet, address, formula text, cached value/display text, and rough dependency tokens where practical.

#### `excel_list_defined_names(path)`

Returns workbook and sheet-scoped names with formulas/ranges.

#### `excel_extract_vba(path)`

Returns static VBA modules, without executing Excel:

```json
{
  "hasVbaProject": true,
  "modules": [
    {
      "name": "Module1",
      "kind": "standardModule",
      "lineCount": 120,
      "code": "Option Explicit\n..."
    }
  ]
}
```

#### `excel_analyze_vba(path)`

Later layer over `excel_extract_vba`:

- procedures/functions with signatures;
- event handlers (`Workbook_Open`, `Worksheet_Change`, button handlers);
- calls between procedures;
- Excel object model references (`Worksheets(...)`, `Range(...)`, `Cells(...)`);
- file/database/network dependencies;
- conversion hints for C# services/classes.

## Error Model

Reuse current mcpOffice error code style:

- `file_not_found`
- `file_exists`
- `invalid_path`
- `unsupported_format`
- `parse_error`
- `index_out_of_range`
- `io_error`
- `internal_error`

Potential Excel-specific additions, only if needed:

- `sheet_not_found`
- `range_too_large`
- `vba_project_missing`
- `vba_project_locked`
- `vba_parse_error`

## Implementation Plan

1. Add DevExpress Spreadsheet package references.
2. Add Excel DTOs and `IExcelWorkbookService`.
3. Implement `excel_list_sheets`.
4. Implement `excel_read_sheet` with `maxCells`.
5. Add integration test for listing tools and reading a generated workbook.
6. Spike `excel_extract_vba` against `C:\temp\macro\Air - Labware.xlsm`.
7. Decide whether static VBA extraction is implemented in-process via `OpenMcdf` or deferred behind an optional extractor.
8. Implement formula/structure tools after basic sheet reading is stable.

## Out Of Scope For This POC

- Executing macros.
- Modifying macros.
- Full Excel calculation engine parity.
- Preserving every visual formatting detail.
- Converting VBA to C# automatically in one step. The server exposes facts and source code; the agent performs the migration reasoning.
