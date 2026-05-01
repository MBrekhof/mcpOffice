# Test fixtures

Hand-authored binary fixtures used by `mcpOffice.Tests` and `mcpOffice.Tests.Integration`. Most Word and Excel tests generate their fixtures programmatically; the files in this directory exist because the relevant test target can't be authored programmatically (yet).

## sample-with-macros.xlsm

Tiny `.xlsm` for VBA extraction tests. Contains:

- `Module1` (standard module): a `Sub Hello()` that calls `Debug.Print "hi"`.
- `ThisWorkbook` (document module): an empty `Private Sub Workbook_Open()` event handler.

Used by:

- `tests/mcpOffice.Tests/Excel/Vba/VbaProjectReaderTests.Reads_modules_from_real_excel_fixture` — exercises the zip path and Excel's real copy-token compressed chunks (synthetic builder tests use literal-only chunks only).
- `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.Extract_vba_via_stdio_returns_modules` — stdio integration test for `excel_extract_vba`.

### Regenerating

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
