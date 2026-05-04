# Test fixtures

Hand-authored binary fixtures used by `mcpOffice.Tests` and `mcpOffice.Tests.Integration`. Most Word and Excel tests generate their fixtures programmatically; the files in this directory exist because the relevant test target can't be authored programmatically (yet).

## sample-with-macros.xlsm

Tiny `.xlsm` for VBA extraction tests. Contains:

- `Module1` (standard module): a `Sub Hello()` that calls `Debug.Print "hi"`.
- `ThisWorkbook` (document module): an empty `Private Sub Workbook_Open()` event handler.

Used by:

- `tests/mcpOffice.Tests/Excel/Vba/VbaProjectReaderTests.Reads_modules_from_real_excel_fixture` — exercises the zip path and Excel's real copy-token compressed chunks (synthetic builder tests use literal-only chunks only).
- `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.Extract_vba_via_stdio_returns_modules` — stdio integration test for `excel_extract_vba`.

## synthetic-vba.xlsm

Richer `.xlsm` for end-to-end pipeline tests (xlsm → vbaProject.bin → MS-OVBA decompression → analyzer). Contains:

- `Module1` (standard module): `Main`, `Process(ByVal r As Range)`, `Variadic(ParamArray args() As Variant)`, `Static Sub StatefulCount`. Exercises ParamArray and Static-Sub forms in a real Excel-authored project (synthetic `VbaProjectBinBuilder` covers them too, but only via literal-only compressed chunks).
- `ThisWorkbook` (document module): `Private Sub Workbook_Open()` that calls `Main` — load-bearing cross-module call edge for the call-graph assertion.
- `Blad1` / `Sheet1` (document module, codename is locale-dependent): `Private Sub Worksheet_Change(ByVal Target As Range)`.
- `Class1` (class module): `Public Sub Greet(ByVal who As String)`.

Used by:

- `tests/mcpOffice.Tests/Excel/Vba/SyntheticAnalyzeTests.cs` — unconditional end-to-end pipeline test, complements the gated `AirSampleAnalysisTests` so suites running on machines without `C:\Projects\mcpOffice-samples\Air.xlsm` still get full-pipeline coverage.

### Regenerating

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/fixtures/Generate-SyntheticVbaXlsm.ps1
```

Requires Excel installed and "Trust access to the VBA project object model" enabled (File → Options → Trust Center → Trust Center Settings → Macro Settings). On Dutch Excel: "Vertrouwen geven aan toegang tot het VBA-projectobjectmodel".

The script discovers the workbook/sheet codenames at runtime via VBComponents enumeration (locale-independent — works on Dutch, English, etc.).

## Regenerating sample-with-macros.xlsm

DevExpress.Spreadsheet cannot author VBA, so this fixture is created manually:

1. Open Excel → New blank workbook.
2. Alt+F11 → Insert → Module → name it `Module1`. Body:
   ```vb
   Sub Hello()
     Debug.Print "hi"
   End Sub
   ```
3. Project Explorer → double-click `ThisWorkbook`. Body:
   ```vb
   Private Sub Workbook_Open()
   End Sub
   ```
4. **File → Save As**. In the file-type dropdown, **explicitly select "Excel Macro-Enabled Workbook (*.xlsm)"** — typing `.xlsm` in the filename is not enough; Excel uses the dropdown to decide whether to keep the VBA project.
5. Save into `tests/fixtures/sample-with-macros.xlsm`. Target size <30 KB.

To verify the file has the VBA project, peek inside as a ZIP and confirm `xl/vbaProject.bin` is present.
