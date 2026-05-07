# `excel_export_csv` Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Ship `excel_export_csv` (the 27th MCP tool) — streams a worksheet or A1 range to a CSV file on disk for `pandas.read_csv` / `polars.read_csv` consumption.

**Architecture:** New streaming `CsvWriter` (RFC 4180 quoting, invariant-culture formatting, ISO 8601 datetimes) lives under `Services/Excel/Csv/`. New service method `ExportCsv` on `ExcelWorkbookService` reuses existing `LoadWorkbook` / `ResolveWorksheet` / `GetCellValue` primitives, walks the resolved range, and feeds rows to `CsvWriter`. New tool method on `ExcelTools` is a one-line delegate.

**Tech Stack:** .NET 9 · DevExpress.Document.Processor · ModelContextProtocol C# SDK · xUnit (no FluentAssertions in Excel tests — match local style).

**Reference design:** `docs/plans/2026-05-07-mcpoffice-excel-export-csv-design.md`. Read it first.

---

## Conventions used in this plan

- Paths relative to repo root `C:\Projects\mcpOffice\`.
- "Run unit tests" = `dotnet test tests/mcpOffice.Tests --nologo --logger "console;verbosity=normal"`.
- "Run all tests" = `dotnet test --nologo`.
- TDD cycle: write failing test → run (verify red) → minimal implementation → run (verify green) → commit. Each task spells out the cycle.
- Conventional Commits: `feat:`, `test:`, `chore:`, `docs:`.
- After every task: `dotnet build` is 0 warnings / 0 errors AND every prior test still passes. If either breaks, stop and fix before the next task (per superpowers:verification-before-completion).
- Branch off `main` with `git checkout -b feat/excel-export-csv` before Task 1.

---

# Phase 0 — Branch + DTO

### Task 1: Branch + add `ExcelExportCsvResult` record

**Files:**
- Create: `src/mcpOffice/Models/ExcelExportCsvResult.cs`

**Step 1: Branch**
```bash
git checkout main
git pull --ff-only
git checkout -b feat/excel-export-csv
```

**Step 2: Write the record**

```csharp
// src/mcpOffice/Models/ExcelExportCsvResult.cs
namespace McpOffice.Models;

public sealed record ExcelExportCsvResult(
    string OutputPath,
    int RowCount,
    int ColumnCount,
    long BytesWritten);
```

No test — pure DTO. The first consumer (the service) will exercise its construction.

**Step 3: Build**
```bash
dotnet build --nologo
```
Expected: 0 warnings, 0 errors.

**Step 4: Commit**
```bash
git add src/mcpOffice/Models/ExcelExportCsvResult.cs
git commit -m "feat: add ExcelExportCsvResult DTO"
```

---

# Phase 1 — CsvWriter

The writer is a thin facade over a `TextWriter`: writes one `IReadOnlyList<object?>` row at a time, handles RFC 4180 quoting, formats values via invariant culture. No workbook coupling; tests use `MemoryStream` only.

### Task 2: CsvWriter — text rows + RFC 4180 quoting

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Csv/CsvWriterTests.cs`
- Create: `src/mcpOffice/Services/Excel/Csv/CsvWriter.cs`

**Step 1: Write the failing tests**

```csharp
// tests/mcpOffice.Tests/Excel/Csv/CsvWriterTests.cs
using System.Text;
using McpOffice.Services.Excel.Csv;

namespace McpOffice.Tests.Excel.Csv;

public class CsvWriterTests
{
    [Fact]
    public void Writes_plain_text_rows_with_crlf_between_rows_and_no_trailing_crlf()
    {
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { "a", "b", "c" });
            writer.WriteRow(new object?[] { "d", "e", "f" });
        });

        var text = Encoding.UTF8.GetString(bytes);
        Assert.Equal("a,b,c\r\nd,e,f", text);
    }

    [Fact]
    public void Quotes_text_containing_comma_quote_or_newline()
    {
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { "no special", "has,comma", "has\"quote", "has\nnewline", "has\rcr" });
        });

        var text = Encoding.UTF8.GetString(bytes);
        Assert.Equal(
            "no special,\"has,comma\",\"has\"\"quote\",\"has\nnewline\",\"has\rcr\"",
            text);
    }

    private static byte[] WriteToBytes(Action<CsvWriter> write)
    {
        using var stream = new MemoryStream();
        using (var writer = new CsvWriter(stream))
        {
            write(writer);
        }
        return stream.ToArray();
    }
}
```

**Step 2: Run — fails (`CsvWriter` doesn't exist)**
```bash
dotnet test tests/mcpOffice.Tests --filter "FullyQualifiedName~CsvWriterTests" --nologo
```
Expected: compile error / test failure.

**Step 3: Implement minimal CsvWriter**

```csharp
// src/mcpOffice/Services/Excel/Csv/CsvWriter.cs
using System.Globalization;
using System.Text;

namespace McpOffice.Services.Excel.Csv;

public sealed class CsvWriter : IDisposable
{
    // UTF-8 without BOM. pandas.read_csv default; BOM breaks naive consumers.
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);
    private const string LineSeparator = "\r\n";

    private readonly StreamWriter _writer;
    private bool _firstRow = true;

    public CsvWriter(Stream stream)
    {
        _writer = new StreamWriter(stream, Utf8NoBom, bufferSize: 64 * 1024, leaveOpen: false);
    }

    public void WriteRow(IReadOnlyList<object?> values)
    {
        if (!_firstRow) _writer.Write(LineSeparator);
        _firstRow = false;

        for (var i = 0; i < values.Count; i++)
        {
            if (i > 0) _writer.Write(',');
            _writer.Write(FormatField(values[i]));
        }
    }

    private static string FormatField(object? value)
    {
        if (value is null) return string.Empty;
        var text = value.ToString() ?? string.Empty;
        return Quote(text);
    }

    private static string Quote(string text)
    {
        if (text.Length == 0) return text;
        if (text.IndexOfAny(['"', ',', '\r', '\n']) < 0) return text;
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }

    public void Dispose() => _writer.Dispose();
}
```

**Step 4: Run — passes**
```bash
dotnet test tests/mcpOffice.Tests --filter "FullyQualifiedName~CsvWriterTests" --nologo
```

**Step 5: Commit**
```bash
git add tests/mcpOffice.Tests/Excel/Csv/CsvWriterTests.cs src/mcpOffice/Services/Excel/Csv/CsvWriter.cs
git commit -m "feat: CsvWriter with RFC 4180 quoting"
```

---

### Task 3: CsvWriter — typed value formatting

Adds invariant-culture number formatting, ISO 8601 datetime, lowercase boolean, null-as-empty handling. Pure formatting tests against the writer; no workbook involvement.

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/Csv/CsvWriterTests.cs`
- Modify: `src/mcpOffice/Services/Excel/Csv/CsvWriter.cs`

**Step 1: Add failing tests**

Append to `CsvWriterTests`:

```csharp
[Fact]
public void Numbers_use_invariant_culture_with_no_thousand_separators()
{
    var bytes = WriteToBytes(writer =>
    {
        writer.WriteRow(new object?[] { 0.21, 1234567.89, 42, -7.5, 0.0 });
    });

    Assert.Equal("0.21,1234567.89,42,-7.5,0", Encoding.UTF8.GetString(bytes));
}

[Fact]
public void DateTime_uses_iso_8601_with_T_separator_and_seconds()
{
    var dt = new DateTime(2026, 5, 7, 14, 30, 0, DateTimeKind.Unspecified);
    var midnight = new DateTime(2026, 5, 7, 0, 0, 0, DateTimeKind.Unspecified);

    var bytes = WriteToBytes(writer =>
    {
        writer.WriteRow(new object?[] { dt, midnight });
    });

    Assert.Equal("2026-05-07T14:30:00,2026-05-07T00:00:00", Encoding.UTF8.GetString(bytes));
}

[Fact]
public void Boolean_emits_lowercase()
{
    var bytes = WriteToBytes(writer =>
    {
        writer.WriteRow(new object?[] { true, false });
    });

    Assert.Equal("true,false", Encoding.UTF8.GetString(bytes));
}

[Fact]
public void Null_and_empty_string_both_emit_empty_unquoted_field()
{
    var bytes = WriteToBytes(writer =>
    {
        writer.WriteRow(new object?[] { null, "", "x" });
    });

    Assert.Equal(",,x", Encoding.UTF8.GetString(bytes));
}

[Fact]
public void NlNL_host_culture_does_not_leak_into_output()
{
    var original = System.Globalization.CultureInfo.CurrentCulture;
    try
    {
        System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("nl-NL");
        var bytes = WriteToBytes(writer =>
        {
            writer.WriteRow(new object?[] { 0.21, 1234.5 });
        });
        Assert.Equal("0.21,1234.5", Encoding.UTF8.GetString(bytes));
    }
    finally
    {
        System.Globalization.CultureInfo.CurrentCulture = original;
    }
}
```

**Step 2: Run — fails (numbers/dates currently round-trip via `value.ToString()` which honours `CurrentCulture`)**

**Step 3: Implement** — replace `FormatField` in `CsvWriter`:

```csharp
private static string FormatField(object? value)
{
    return value switch
    {
        null                 => string.Empty,
        bool b               => b ? "true" : "false",
        DateTime dt          => dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
        DateTimeOffset dto   => dto.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
        IFormattable f       => Quote(f.ToString(format: null, formatProvider: CultureInfo.InvariantCulture)),
        string s             => Quote(s),
        _                    => Quote(value.ToString() ?? string.Empty),
    };
}
```

`IFormattable` covers `double`, `decimal`, `int`, `long`, etc. — all numerics flow through invariant culture. Booleans and DateTimes are handled before the `IFormattable` arm (DateTime is `IFormattable` but its default `g` format is locale-dependent, so we intercept first).

> Note: numbers from `IFormattable` are routed through `Quote`, but the default invariant `ToString()` for built-in numerics never produces `,` `\r` `\n` `"`, so they pass through unquoted. Cheap belt-and-braces.

**Step 4: Run — passes**

**Step 5: Commit**
```bash
git add tests/mcpOffice.Tests/Excel/Csv/CsvWriterTests.cs src/mcpOffice/Services/Excel/Csv/CsvWriter.cs
git commit -m "feat: CsvWriter formats values via invariant culture"
```

---

### Task 4: CsvWriter — UTF-8 no-BOM byte-level confirmation

A single test that locks the encoding contract. Cheap insurance against a future refactor that swaps `StreamWriter`'s encoding.

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/Csv/CsvWriterTests.cs`

**Step 1: Add test**

```csharp
[Fact]
public void Output_is_utf8_without_bom()
{
    var bytes = WriteToBytes(writer =>
    {
        writer.WriteRow(new object?[] { "héllo" });
    });

    // No UTF-8 BOM (EF BB BF) at the start.
    Assert.NotEqual(0xEF, bytes[0]);

    // Round-trips as UTF-8.
    Assert.Equal("héllo", Encoding.UTF8.GetString(bytes));
    Assert.Equal(new byte[] { 0x68, 0xC3, 0xA9, 0x6C, 0x6C, 0x6F }, bytes);
}
```

**Step 2: Run — passes (no impl change needed; the `Utf8NoBom` constant is already correct).**

**Step 3: Commit**
```bash
git add tests/mcpOffice.Tests/Excel/Csv/CsvWriterTests.cs
git commit -m "test: lock CsvWriter UTF-8 no-BOM contract"
```

---

# Phase 2 — Service method

### Task 5: `ExportCsv` on `IExcelWorkbookService` + happy path

Adds the interface signature, the implementation, and one happy-path test that exercises the full pipeline (load → resolve → walk → write).

**Files:**
- Modify: `src/mcpOffice/Services/Excel/IExcelWorkbookService.cs`
- Modify: `src/mcpOffice/Services/Excel/ExcelWorkbookService.cs`
- Create: `tests/mcpOffice.Tests/Excel/ExportCsvTests.cs`

**Step 1: Write the failing test**

```csharp
// tests/mcpOffice.Tests/Excel/ExportCsvTests.cs
using DevExpress.Spreadsheet;
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel;

public class ExportCsvTests
{
    [Fact]
    public void Writes_used_range_with_mixed_types()
    {
        var input = TestExcelWorkbooks.Create(workbook =>
        {
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["C1"].Value = "Active";
            sheet.Cells["D1"].Value = "Joined";
            sheet.Cells["A2"].Value = "Ada";
            sheet.Cells["B2"].Value = 0.21;
            sheet.Cells["C2"].Value = true;
            sheet.Cells["D2"].Value = new DateTime(2026, 5, 7, 14, 30, 0);
            sheet.Cells["A3"].Value = "Bob, Jr.";   // forces quoting
            sheet.Cells["B3"].Value = 1234567.89;
            sheet.Cells["C3"].Value = false;
            // D3 left blank
        });
        var output = TempPath(".csv");

        try
        {
            var result = new ExcelWorkbookService().ExportCsv(
                input,
                output,
                sheetName: "Data",
                sheetIndex: null,
                range: null,
                overwrite: false,
                maxRows: 1_048_576);

            Assert.Equal(output, result.OutputPath);
            Assert.Equal(3, result.RowCount);
            Assert.Equal(4, result.ColumnCount);
            Assert.True(result.BytesWritten > 0);

            var text = File.ReadAllText(output);
            Assert.Equal(
                "Name,Amount,Active,Joined\r\n" +
                "Ada,0.21,true,2026-05-07T14:30:00\r\n" +
                "\"Bob, Jr.\",1234567.89,false,",
                text);
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    private static string TempPath(string ext) =>
        Path.Combine(Path.GetTempPath(), $"mcpoffice-csv-{Guid.NewGuid():N}{ext}");
}
```

**Step 2: Run — fails (no method on service)**

**Step 3: Extend interface**

```csharp
// src/mcpOffice/Services/Excel/IExcelWorkbookService.cs — append to interface
ExcelExportCsvResult ExportCsv(
    string path,
    string outputPath,
    string? sheetName,
    int? sheetIndex,
    string? range,
    bool overwrite,
    int maxRows);
```

**Step 4: Implement on the service**

Add to `ExcelWorkbookService`:

```csharp
using McpOffice.Services.Excel.Csv;

// ... inside the class:

public ExcelExportCsvResult ExportCsv(
    string path,
    string outputPath,
    string? sheetName,
    int? sheetIndex,
    string? range,
    bool overwrite,
    int maxRows)
{
    PathGuard.RequireExists(path);
    PathGuard.RequireWritable(outputPath, overwrite);

    try
    {
        using var workbook = LoadWorkbook(path);
        var worksheet = ResolveWorksheet(workbook, sheetName, sheetIndex);
        var cellRange = string.IsNullOrWhiteSpace(range)
            ? worksheet.GetUsedRange()
            : worksheet.Range[range];

        var rangeReference = cellRange.GetReferenceA1();
        if (cellRange.RowCount > maxRows)
        {
            throw ToolError.RangeTooLarge(rangeReference, cellRange.RowCount, maxRows);
        }

        long bytesWritten;
        using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
        using (var csv = new CsvWriter(fileStream))
        {
            for (var r = 0; r < cellRange.RowCount; r++)
            {
                var row = new object?[cellRange.ColumnCount];
                for (var c = 0; c < cellRange.ColumnCount; c++)
                {
                    row[c] = GetCellValue(cellRange[r, c].Value);
                }
                csv.WriteRow(row);
            }
            csv.Dispose();
            bytesWritten = fileStream.Length;
        }

        return new ExcelExportCsvResult(
            outputPath,
            cellRange.RowCount,
            cellRange.ColumnCount,
            bytesWritten);
    }
    catch (Exception ex) when (ex is not McpException)
    {
        throw ToolError.ParseError(path, ex.Message);
    }
}
```

> The `RequireWritable` call deliberately runs **before** the workbook load. Failing fast on `file_exists` / `invalid_path` for `outputPath` saves an unnecessary workbook open on misuse. The order also means `RangeTooLarge` will not delete a pre-existing output file. (`FileMode.Create` opens-or-truncates, but throws on read-only — the precondition surfaces a clean error first.)

> `RangeTooLarge`'s message says "exceeds maxCells=Y" — the service passes `maxRows` into the same constructor. The wording is slightly off (rows vs cells); the design accepts this. If it irritates a real consumer, add a `RangeTooLargeRows(...)` helper later.

**Step 5: Run — passes**
```bash
dotnet test tests/mcpOffice.Tests --filter "FullyQualifiedName~ExportCsvTests" --nologo
```

**Step 6: Commit**
```bash
git add src/mcpOffice/Services/Excel/IExcelWorkbookService.cs src/mcpOffice/Services/Excel/ExcelWorkbookService.cs tests/mcpOffice.Tests/Excel/ExportCsvTests.cs
git commit -m "feat: ExportCsv service method + happy-path test"
```

---

### Task 6: Range slicing

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/ExportCsvTests.cs`

**Step 1: Add failing test**

```csharp
[Fact]
public void Range_parameter_slices_to_subset()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        var sheet = workbook.Worksheets[0];
        for (var r = 1; r <= 5; r++)
        {
            for (var c = 0; c < 4; c++)
            {
                sheet.Cells[r - 1, c].Value = $"r{r}c{c + 1}";
            }
        }
    });
    var output = TempPath(".csv");

    try
    {
        var result = new ExcelWorkbookService().ExportCsv(
            input, output, null, null, range: "A1:B3", overwrite: false, maxRows: 1_048_576);

        Assert.Equal(3, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(
            "r1c1,r1c2\r\nr2c1,r2c2\r\nr3c1,r3c2",
            File.ReadAllText(output));
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}
```

**Step 2: Run — passes (range plumbing already in service from Task 5).**

**Step 3: Commit**
```bash
git add tests/mcpOffice.Tests/Excel/ExportCsvTests.cs
git commit -m "test: ExportCsv range slicing"
```

---

### Task 7: Sheet resolution variants

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/ExportCsvTests.cs`

**Step 1: Add failing tests**

```csharp
[Fact]
public void Resolves_by_sheet_index_when_name_is_omitted()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        workbook.Worksheets[0].Name = "First";
        workbook.Worksheets[0].Cells["A1"].Value = "one";
        workbook.Worksheets.Add("Second");
        workbook.Worksheets[1].Cells["A1"].Value = "two";
    });
    var output = TempPath(".csv");

    try
    {
        var result = new ExcelWorkbookService().ExportCsv(
            input, output, null, sheetIndex: 1, range: null, overwrite: false, maxRows: 1_048_576);

        Assert.Equal("two", File.ReadAllText(output));
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}

[Fact]
public void Throws_sheet_not_found_for_unknown_name()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        workbook.Worksheets[0].Cells["A1"].Value = "x";
    });
    var output = TempPath(".csv");

    try
    {
        var ex = Assert.Throws<ModelContextProtocol.McpException>(() =>
            new ExcelWorkbookService().ExportCsv(
                input, output, sheetName: "Missing", sheetIndex: null, range: null,
                overwrite: false, maxRows: 1_048_576));

        Assert.Contains("[sheet_not_found]", ex.Message);
        Assert.Contains("Missing", ex.Message);
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}
```

**Step 2: Run — passes (existing `ResolveWorksheet` covers both paths).**

**Step 3: Commit**
```bash
git add tests/mcpOffice.Tests/Excel/ExportCsvTests.cs
git commit -m "test: ExportCsv sheet resolution variants"
```

---

### Task 8: Overwrite + output-directory creation

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/ExportCsvTests.cs`

**Step 1: Add failing tests**

```csharp
[Fact]
public void Overwrite_false_throws_file_exists_when_output_exists()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        workbook.Worksheets[0].Cells["A1"].Value = "x";
    });
    var output = TempPath(".csv");
    File.WriteAllText(output, "preexisting");

    try
    {
        var ex = Assert.Throws<ModelContextProtocol.McpException>(() =>
            new ExcelWorkbookService().ExportCsv(
                input, output, null, null, null, overwrite: false, maxRows: 1_048_576));

        Assert.Contains("[file_exists]", ex.Message);
        Assert.Equal("preexisting", File.ReadAllText(output));
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}

[Fact]
public void Overwrite_true_replaces_existing_file()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        workbook.Worksheets[0].Cells["A1"].Value = "fresh";
    });
    var output = TempPath(".csv");
    File.WriteAllText(output, "stale");

    try
    {
        new ExcelWorkbookService().ExportCsv(
            input, output, null, null, null, overwrite: true, maxRows: 1_048_576);

        Assert.Equal("fresh", File.ReadAllText(output));
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}

[Fact]
public void Creates_missing_output_directory()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        workbook.Worksheets[0].Cells["A1"].Value = "x";
    });
    var subdir = Path.Combine(Path.GetTempPath(), $"mcpoffice-csv-{Guid.NewGuid():N}", "nested");
    var output = Path.Combine(subdir, "out.csv");

    try
    {
        new ExcelWorkbookService().ExportCsv(
            input, output, null, null, null, overwrite: false, maxRows: 1_048_576);

        Assert.True(File.Exists(output));
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (Directory.Exists(subdir)) Directory.Delete(subdir, recursive: true);
    }
}
```

**Step 2: Run — passes (`PathGuard.RequireWritable` already does both: errors on `file_exists` and creates the parent directory).**

**Step 3: Commit**
```bash
git add tests/mcpOffice.Tests/Excel/ExportCsvTests.cs
git commit -m "test: ExportCsv overwrite semantics + dir creation"
```

---

### Task 9: `maxRows` guard

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/ExportCsvTests.cs`

**Step 1: Add failing test**

```csharp
[Fact]
public void Throws_range_too_large_when_row_count_exceeds_max_rows()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        var sheet = workbook.Worksheets[0];
        for (var r = 1; r <= 10; r++)
        {
            sheet.Cells[r - 1, 0].Value = $"row{r}";
        }
    });
    var output = TempPath(".csv");

    try
    {
        var ex = Assert.Throws<ModelContextProtocol.McpException>(() =>
            new ExcelWorkbookService().ExportCsv(
                input, output, null, null, null, overwrite: false, maxRows: 5));

        Assert.Contains("[range_too_large]", ex.Message);
        Assert.False(File.Exists(output), "no output file should be created when maxRows is exceeded");
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}
```

**Step 2: Run — passes (guard added in Task 5).**

**Step 3: Commit**
```bash
git add tests/mcpOffice.Tests/Excel/ExportCsvTests.cs
git commit -m "test: ExportCsv maxRows guard"
```

---

### Task 10: Formula cells, empty cells, locale-neutral output

Three small tests that lock the value-extraction contract.

**Files:**
- Modify: `tests/mcpOffice.Tests/Excel/ExportCsvTests.cs`

**Step 1: Add failing tests**

```csharp
[Fact]
public void Formula_cell_emits_cached_value_not_formula_text()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 40;
        sheet.Cells["B1"].Formula = "=A1+2";
        workbook.Calculate();
    });
    var output = TempPath(".csv");

    try
    {
        new ExcelWorkbookService().ExportCsv(
            input, output, null, null, null, overwrite: false, maxRows: 1_048_576);

        Assert.Equal("40,42", File.ReadAllText(output));
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}

[Fact]
public void Empty_cells_emit_empty_unquoted_fields()
{
    var input = TestExcelWorkbooks.Create(workbook =>
    {
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "x";
        // B1, C1 left blank
        sheet.Cells["D1"].Value = "y";
    });
    var output = TempPath(".csv");

    try
    {
        new ExcelWorkbookService().ExportCsv(
            input, output, null, null, range: "A1:D1", overwrite: false, maxRows: 1_048_576);

        Assert.Equal("x,,,y", File.ReadAllText(output));
    }
    finally
    {
        if (File.Exists(input)) File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}

[Fact]
public void Output_is_locale_neutral_when_host_culture_is_nlNL()
{
    var original = System.Globalization.CultureInfo.CurrentCulture;
    try
    {
        System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("nl-NL");

        var input = TestExcelWorkbooks.Create(workbook =>
        {
            workbook.Worksheets[0].Cells["A1"].Value = 0.21;
            workbook.Worksheets[0].Cells["B1"].Value = 1234567.89;
        });
        var output = TempPath(".csv");

        try
        {
            new ExcelWorkbookService().ExportCsv(
                input, output, null, null, null, overwrite: false, maxRows: 1_048_576);

            Assert.Equal("0.21,1234567.89", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }
    finally
    {
        System.Globalization.CultureInfo.CurrentCulture = original;
    }
}
```

**Step 2: Run — passes (CsvWriter formats via invariant culture; `LoadWorkbook` already pins workbook culture; `GetCellValue` returns `null` for empty cells which the writer maps to `string.Empty`).**

**Step 3: Commit**
```bash
git add tests/mcpOffice.Tests/Excel/ExportCsvTests.cs
git commit -m "test: ExportCsv formulas, empty cells, locale-neutral output"
```

---

# Phase 3 — Tool surface

### Task 11: Add `excel_export_csv` tool method

**Files:**
- Modify: `src/mcpOffice/Tools/ExcelTools.cs`

**Step 1: Append the tool method** (place after `ExcelReadSheet`, alphabetical-ish grouping isn't enforced):

```csharp
[McpServerTool(Name = "excel_export_csv")]
[Description("Streams a worksheet (or A1 range) to a CSV file on disk for pandas/polars consumption. RFC 4180 dialect, UTF-8 (no BOM), CRLF line endings, invariant-culture numbers, ISO 8601 datetimes, lowercase booleans. Formula cells emit their cached value (no formula text). Returns {outputPath, rowCount, columnCount, bytesWritten}.")]
public static object ExcelExportCsv(
    [Description("Absolute path to the .xlsx/.xlsm input workbook")] string path,
    [Description("Absolute path to the .csv output file. Parent directory is created if missing.")] string outputPath,
    [Description("Worksheet name. If omitted, sheetIndex is used.")] string? sheetName = null,
    [Description("0-based worksheet index used when sheetName is omitted. Defaults to 0.")] int? sheetIndex = null,
    [Description("Optional A1 range such as A1:D20. Defaults to the worksheet used range.")] string? range = null,
    [Description("Overwrite outputPath if it already exists. Defaults to false.")] bool overwrite = false,
    [Description("Maximum rows to export. Defaults to 1,048,576 (Excel row ceiling). Trips range_too_large if exceeded.")] int maxRows = 1_048_576)
    => Service.ExportCsv(path, outputPath, sheetName, sheetIndex, range, overwrite, maxRows);
```

**Step 2: Build**
```bash
dotnet build --nologo
```
Expected: 0 warnings, 0 errors.

**Step 3: Run all unit tests** to confirm nothing broke
```bash
dotnet test tests/mcpOffice.Tests --nologo
```

**Step 4: Commit**
```bash
git add src/mcpOffice/Tools/ExcelTools.cs
git commit -m "feat: excel_export_csv tool"
```

---

### Task 12: Update `ToolSurfaceTests` to expect 27 tools

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`

**Step 1: Insert `"excel_export_csv"` into the expected array** (after `"excel_extract_vba"`, alphabetical):

```csharp
"excel_analyze_vba",
"excel_export_csv",
"excel_extract_vba",
"excel_get_metadata",
// ...
```

**Step 2: Run integration tests**
```bash
dotnet test tests/mcpOffice.Tests.Integration --nologo
```
Expected: passes — server now exposes 27 tools, the catalog test asserts the new one is present.

**Step 3: Commit**
```bash
git add tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs
git commit -m "test: ToolSurfaceTests covers excel_export_csv"
```

---

### Task 13: End-to-end stdio integration test

One happy-path round-trip through the JSON-RPC layer. Mirror the shape of existing `ExcelWorkflowTests.Read_sheet_via_stdio`.

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs`

**Step 1: Add the test** (append to the class):

```csharp
[Fact]
public async Task Export_csv_via_stdio()
{
    var input = TempPath(".xlsx");
    var output = TempPath(".csv");
    try
    {
        using (var workbook = new Workbook())
        {
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Ada";
            sheet.Cells["B2"].Value = 0.21;
            workbook.SaveDocument(input, SpreadsheetFormat.Xlsx);
        }

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_export_csv",
            new Dictionary<string, object?>
            {
                ["path"] = input,
                ["outputPath"] = output,
                ["sheetName"] = "Data",
            });
        var text = result.Content.OfType<TextContentBlock>().Single().Text;

        Assert.Contains("\"rowCount\":2", text);
        Assert.Contains("\"columnCount\":2", text);
        Assert.True(File.Exists(output));
        Assert.Equal("Name,Amount\r\nAda,0.21", File.ReadAllText(output));
    }
    finally
    {
        if (File.Exists(input))  File.Delete(input);
        if (File.Exists(output)) File.Delete(output);
    }
}
```

> The class likely already has a `TempPath(string)` helper from the existing tests. If it doesn't, add the same one used in `ExportCsvTests`.

**Step 2: Run integration tests**
```bash
dotnet test tests/mcpOffice.Tests.Integration --nologo
```

**Step 3: Commit**
```bash
git add tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs
git commit -m "test: stdio round-trip for excel_export_csv"
```

---

# Phase 4 — Verification + housekeeping

### Task 14: Final verify — clean build, all green

**Step 1: Release build + full test pass**
```bash
dotnet build -c Release --nologo
dotnet test  -c Release --nologo
```
Expected: 0 warnings, 0 errors. Every prior test still passes; new tests from this branch pass. Skipped count unchanged (Air.xlsm gated test still skips).

If anything fails, stop and fix before continuing.

### Task 15: Live verification against `Air.xlsm` (optional — gated)

Per the design doc's "Live verification" section. Skip if `C:\Projects\mcpOffice-samples\Air.xlsm` is absent.

**Step 1: Run a one-off PowerShell snippet** that calls the published server via the same harness style as the integration tests, exporting one sheet to a temp CSV. Quick approach: write a throwaway xUnit fact under `[Fact(Skip = "manual verification")]` in `ExportCsvTests.cs`, unskip locally, run, re-skip. Or use `dotnet run` + Claude Code's MCP wiring.

**Step 2: Confirm with a one-line Python check**
```powershell
python -c "import pandas as pd; df = pd.read_csv(r'C:\\temp\\air-export.csv'); print(df.shape, df.dtypes)"
```
Expected: shape matches `excel_list_sheets`'s reported `rowCount` × `columnCount` for that sheet, and dtypes look sensible (numbers as `float64` / `int64`, dates parsed if `parse_dates=` passed).

This step is informational — no code changes. If the dtypes look wrong (e.g. dates landing as `object`), that's a `pandas` type-coercion question, not a CSV bug.

### Task 16: Update `TODO.md` and `SESSION_HANDOFF.md`

**Files:**
- Modify: `TODO.md`
- Modify: `SESSION_HANDOFF.md`

**Step 1: TODO.md**

Move the `excel_export_csv` line from the "Carried from Excel POC" section into the analyzer-style "DONE" header pattern used elsewhere. Leave the NDJSON sibling and `.csv.gz` follow-ups as separate bullets under a new "Deferred follow-ups" sub-heading. Match the wording of the v3 deferred-follow-ups list for consistency.

**Step 2: SESSION_HANDOFF.md**

Replace the body with a fresh session handoff: branch state, latest commit hash, build/tests counts (271 unit grows by ~12 tests; 14 integration grows by 1; tool surface grows from 26 → 27), what landed, what's next.

**Step 3: Commit**
```bash
git add TODO.md SESSION_HANDOFF.md
git commit -m "docs: update TODO + SESSION_HANDOFF for excel_export_csv"
```

### Task 17: PR preparation (do not push without user confirmation)

**Step 1:** Confirm branch is clean and ahead of `main`:
```bash
git status
git log --oneline main..feat/excel-export-csv
```

**Step 2:** Surface the branch state to the user and ask whether to:
- Squash-merge locally to `main` and delete the branch (matches the v3 / md-converter pattern), or
- Push to `origin` and open a PR via `gh pr create`.

**Do not push or merge without explicit user direction.** Per global CLAUDE.md, push may hang and merging is destructive on shared state.

---

## What this plan deliberately does NOT do

- **No `.csv.gz` compression.** Trivial follow-up — wrap the `FileStream` in `GZipStream` based on output extension. Adds a test matrix that's not justified for v1.
- **No `excel_export_ndjson` sibling.** Shares streaming infrastructure but is a separate tool; ships next.
- **No CI workflow change.** The repo doesn't have CI yet; verification stays manual `dotnet build` + `dotnet test`.
- **No new error codes.** All paths covered by `file_not_found`, `invalid_path`, `file_exists`, `index_out_of_range`, `sheet_not_found`, `parse_error`, `range_too_large`, `io_error`.
- **No formula recalculation.** Cached value only. `excel_list_formulas` covers the recalc path.
- **No streaming optimisation beyond `StreamWriter`'s 64 KB buffer.** A 1M-row × 50-col workbook is ~50M cell reads — tolerable. If a real consumer hits a wall, profile then optimise.

## Risks called out

1. **DevExpress `cellRange[r, c]` performance on huge ranges.** Indexer access per cell is O(1) but adds up. The Air.xlsm benchmark on `excel_read_sheet`-style scans is fast enough for 5M cells. If a user complains about a 10M+ sheet, switch to `cellRange.GetSubrange(...).Cells` enumeration.
2. **`RangeTooLarge`'s message says "cells" not "rows".** Cosmetic; the design doc accepts it. Add a row-flavoured helper if a consumer is confused.
3. **`PathGuard.RequireWritable` runs before workbook load.** Means a bad input path with a valid output path errors as `file_not_found` (workbook), not `file_exists` (output). Existing convention — matches `word_create_blank` / `word_mail_merge`.
4. **DateTime emits seconds always.** Even for date-only cells. The agent can `df['col'] = df['col'].dt.normalize()` if they want date-only. No producer-side knob.
