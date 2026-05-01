# mcpOffice Excel VBA Extraction Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Do not start implementation until the user greenlights this plan.

**Date:** 2026-05-01
**Branch:** `poc/excel-tools` (continues the Excel POC)
**Goal:** Ship `excel_extract_vba` — a static, in-process VBA source extractor that reads `.xlsm` files without launching Excel, executing macros, or shelling out to Python.

**Reference design:** `docs/plans/2026-05-01-mcpoffice-excel-poc-design.md` — section "VBA Extraction" and the `excel_extract_vba` JSON shape are authoritative.
**Spike findings:** `tests/mcpOffice.Tests/Spikes/VbaExtractionSpike.cs` and `SESSION_HANDOFF.md` (`## VBA Extraction Spike — Findings`). The spike already proved viability; this plan promotes it to production code with proper tests.

---

## Decisions inherited from the spike

1. **In-process, OpenMcdf-based.** No `olevba`, no Excel COM. The spike confirmed full module discovery + source decompression on a 1.17 MB real-world `vbaProject.bin`.
2. **Production reader lives inside `src/mcpOffice`** under `Services/Excel/Vba/`. No separate library — defer extraction until PowerPoint or another consumer needs it.
3. **`PROJECTVERSION` (id `0x0009`) record special-case.** Size field is hardcoded to `4` but actual payload is `6` bytes (UInt32 major + UInt16 minor). Spike has the working logic; production parser must replicate.
4. **MBCS records first; Unicode siblings preferred when present.** `MODULENAME` (`0x0019`) and `MODULESTREAMNAME` (`0x001A`) are MBCS (cp1252). Their Unicode siblings are `MODULENAMEUNICODE` (`0x0047`) and `MODULESTREAMNAMEUNICODE` (`0x0032`). Production parser prefers the Unicode form when both are present (MS-OVBA specifies unicode siblings as authoritative).
5. **cp1252 needs `System.Text.Encoding.CodePages`** registered once at process start (already true in tests; we'll add a static initializer in the reader).
6. **Hybrid testing strategy.** The reader exposes two entry points: `Read(string xlsmPath)` (zip + OLE + decompress) and `ReadVbaProjectBin(Stream)` (OLE + decompress only). Unit tests construct synthetic `vbaProject.bin` blobs in-memory via a test-only `VbaProjectBinBuilder` and exercise `ReadVbaProjectBin` — fast, parameterizable, no fixture file. A single hand-authored `tests/fixtures/sample-with-macros.xlsm` is committed and used **only** by (a) one "real-Excel-output" smoke test that exercises the zip path and the actual MS-OVBA compressed-chunk decoder against Excel-produced output, and (b) the stdio integration test. This keeps the binary-blob surface area minimal while ensuring we still validate against real Excel output.

---

## Tool surface delta

Adds **one** tool to the existing 18, bringing the catalog to 19. Shape per design doc §`excel_extract_vba`:

```json
{
  "hasVbaProject": true,
  "modules": [
    { "name": "Module1", "kind": "standardModule", "lineCount": 120, "code": "Option Explicit\n..." }
  ]
}
```

`kind` ∈ `{ "standardModule", "classModule", "documentModule" }` mapped from MODULETYPE id `0x0021` → standard, `0x0022` → class/document (further refined via MODULEDOCSTRING / `VB_Base` if needed; spike showed `0x0022` covers `ThisWorkbook` and `Sheet*` document modules — for the POC we collapse to `documentModule` when the module has the host-document `VB_Base` marker, otherwise `classModule`).

---

## New error codes

Add to `src/mcpOffice/ErrorCode.cs` and `ToolError.cs`:

| Code | When |
|---|---|
| `vba_project_missing` | `.xlsx` (or `.xlsm` without `xl/vbaProject.bin`) — agent should not retry. |
| `vba_project_locked` | dir stream is encrypted / project locked for viewing. |
| `vba_parse_error` | OLE walk or MS-OVBA decompression fails on otherwise-valid `vbaProject.bin`. Message includes underlying reason. |

`ParseError` is reserved for the workbook-level path (DevExpress load failure). VBA-specific failures use `vba_parse_error` so agents can branch.

---

## Conventions used in this plan

- Paths relative to `C:\Projects\mcpOffice\`.
- Each task: write failing test → run (red) → implement → run (green) → commit. After every task: `dotnet build` clean, all tests pass. Stop and fix on failure (per superpowers:verification-before-completion).
- "Run tests": `dotnet test --nologo` from repo root.
- Conventional Commits.

---

# Phase 0 — Promote dependencies + add error codes

### Task 1: Move OpenMcdf + cp1252 dependency from tests to server

**Files:**
- Modify: `src/mcpOffice/mcpOffice.csproj` — add `OpenMcdf` 3.1.3 and `System.Text.Encoding.CodePages` 10.0.7.
- Modify: `tests/mcpOffice.Tests/mcpOffice.Tests.csproj` — keep these references (transitive via project ref is fine; explicit refs are also fine — leave as-is to avoid churn).

**Step 1: Add packages**
```bash
dotnet add src/mcpOffice package OpenMcdf --version 3.1.3
dotnet add src/mcpOffice package System.Text.Encoding.CodePages --version 10.0.7
```

**Step 2: Build**
```bash
dotnet build --nologo
```
Expected: 0 warnings, 0 errors.

**Step 3: Commit**
```bash
git commit -am "chore: add OpenMcdf + cp1252 to server project for VBA extraction"
```

---

### Task 2: Add VBA error codes

**Files:**
- Modify: `src/mcpOffice/ErrorCode.cs`
- Modify: `src/mcpOffice/ToolError.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs`

**Step 1: Failing test** (asserts the helpers exist and produce the right `[code]` prefix)
```csharp
// tests/mcpOffice.Tests/Excel/Vba/VbaErrorCodeTests.cs
using FluentAssertions;
using McpOffice;
using ModelContextProtocol;

namespace McpOffice.Tests.Excel.Vba;

public class VbaErrorCodeTests
{
    [Fact]
    public void VbaProjectMissing_has_stable_code()
    {
        var act = () => throw ToolError.VbaProjectMissing(@"C:\book.xlsx");
        act.Should().Throw<McpException>().Which.Message.Should().Contain("vba_project_missing");
    }

    [Fact]
    public void VbaProjectLocked_has_stable_code()
    {
        var act = () => throw ToolError.VbaProjectLocked(@"C:\book.xlsm");
        act.Should().Throw<McpException>().Which.Message.Should().Contain("vba_project_locked");
    }

    [Fact]
    public void VbaParseError_has_stable_code_and_detail()
    {
        var act = () => throw ToolError.VbaParseError(@"C:\book.xlsm", "bad chunk header");
        act.Should().Throw<McpException>()
           .Which.Message.Should().Contain("vba_parse_error").And.Contain("bad chunk header");
    }
}
```

> Note: `mcpOffice.Tests` does not currently reference FluentAssertions. Existing Excel tests use raw `Assert`. To stay consistent with that style, rewrite the asserts using `Assert.Throws<McpException>` + `Assert.Contains` (matching `ListSheetsTests.cs`). FluentAssertions is only on the integration project. Keep the test file but use xUnit asserts.

**Step 2: Implement**
```csharp
// add to src/mcpOffice/ErrorCode.cs
public const string VbaProjectMissing = "vba_project_missing";
public const string VbaProjectLocked = "vba_project_locked";
public const string VbaParseError = "vba_parse_error";
```
```csharp
// add to src/mcpOffice/ToolError.cs
public static Exception VbaProjectMissing(string path) =>
    Throw(ErrorCode.VbaProjectMissing, $"No VBA project in workbook: {path}");

public static Exception VbaProjectLocked(string path) =>
    Throw(ErrorCode.VbaProjectLocked, $"VBA project is locked for viewing: {path}");

public static Exception VbaParseError(string path, string detail) =>
    Throw(ErrorCode.VbaParseError, $"Could not parse VBA project in {path}: {detail}");
```

**Step 3: Run + commit**
```bash
dotnet test tests/mcpOffice.Tests --nologo
git commit -am "feat: add vba_project_missing/_locked/_parse_error error codes"
```

---

# Phase 1 — MS-OVBA decompressor

### Task 3: Promote `MsOvbaDecompressor` from spike to production

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/MsOvbaDecompressor.cs`

**Step 1: Copy implementation verbatim from spike** (`tests/mcpOffice.Tests/Spikes/VbaExtractionSpike.cs` lines 204-288). Adjust namespace to `McpOffice.Services.Excel.Vba`, make the class `internal static` (only consumers are in the same project; no need to expose it).

```csharp
// src/mcpOffice/Services/Excel/Vba/MsOvbaDecompressor.cs
namespace McpOffice.Services.Excel.Vba;

// MS-OVBA 2.4 RLE decompressor for VBA compressed-container streams.
internal static class MsOvbaDecompressor
{
    public static byte[] Decompress(byte[] compressed)
    {
        // ... (verbatim from spike, including CopyTokenLengthBits)
    }
}
```

> The spike's implementation is correct against the real fixture — the `dir` stream of a 107-module workbook decompressed cleanly. Preserve its inline comments referencing MS-OVBA section numbers; they're load-bearing for future maintainers.

**Step 2: Build**
```bash
dotnet build --nologo
```

**Step 3: Commit** — no test yet (Task 4 covers it).
```bash
git commit -am "feat: promote MsOvbaDecompressor from spike to production code"
```

---

### Task 4: Unit tests for `MsOvbaDecompressor`

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Vba/MsOvbaDecompressorTests.cs`

We hand-craft compressed inputs by following MS-OVBA 2.4.1 directly, since:
- Signature byte: `0x01`
- Chunk header (UInt16 LE): `((chunkSize - 3) & 0x0FFF) | (0b011 << 12) | (compressedFlag << 15)`
- Compressed mode: each 8 tokens prefixed by a 1-byte flag mask. Flag bit 0 = literal (1 byte). Flag bit 1 = copy token (UInt16, but for our tests we only need literals).

Tests cover:

1. **Signature missing → throws `InvalidDataException`** (input `[0x00, ...]`).
2. **Single chunk, all literals, ≤ 8 bytes** — input `"ABCDEFGH"` → output `"ABCDEFGH"`.
3. **Two chunks, all literals, > 8 bytes** — round-trip `"abcdefghij"` (10 bytes; exceeds one flag byte).
4. **Copy-token round-trip via known reference vector** — pick a tiny known-good blob from MS-OVBA spec example or from the spike output (commit hex literal as test data) and assert decompression matches expected text.
5. **Bad chunk signature → throws** — flip bits in chunk header so signature ≠ `0b011`.

```csharp
// tests/mcpOffice.Tests/Excel/Vba/MsOvbaDecompressorTests.cs
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class MsOvbaDecompressorTests
{
    [Fact]
    public void Throws_when_signature_byte_missing()
    {
        var bad = new byte[] { 0x00, 0x00, 0x00 };
        Assert.Throws<InvalidDataException>(() => MsOvbaDecompressor.Decompress(bad));
    }

    [Fact]
    public void Decompresses_single_chunk_of_literals()
    {
        // 8 literal bytes "ABCDEFGH": signature + chunk header(compressed=1, size=11) + flag(0x00) + 8 literals
        // chunkSize = 2 header + 1 flag + 8 literals = 11. raw = (11-3) | 0x3000 | 0x8000 = 0xB008.
        var bytes = new byte[]
        {
            0x01,
            0x08, 0xB0,             // chunk header (LE)
            0x00,                    // flag byte: all 8 tokens are literals
            0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48
        };
        var result = MsOvbaDecompressor.Decompress(bytes);
        Assert.Equal("ABCDEFGH", System.Text.Encoding.ASCII.GetString(result));
    }

    [Fact]
    public void Throws_on_bad_chunk_signature()
    {
        // Flip chunk-signature nibble away from 0b011.
        var bytes = new byte[] { 0x01, 0x08, 0x00 /* sig=0 */, 0x00, 0x41 };
        Assert.Throws<InvalidDataException>(() => MsOvbaDecompressor.Decompress(bytes));
    }

    // Test 3 (multi-chunk literals) and Test 4 (copy-token reference vector)
    // follow the same pattern; use spike output to obtain a known-good blob.
}
```

**Step: Run + commit**
```bash
dotnet test tests/mcpOffice.Tests --nologo
git commit -am "test: MsOvbaDecompressor — signature, literals, bad chunk header"
```

> If hand-crafting copy-token tests proves fiddly during implementation, fall back to extracting a small known-good `(compressed, expected)` pair from the spike output and committing it as test data under `tests/mcpOffice.Tests/Excel/Vba/Data/`. Don't write a custom compressor just to test the decompressor — the production code never compresses.

---

# Phase 2 — Dir stream parser

### Task 5: Promote `VbaDirStreamParser` from spike

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaDirStreamParser.cs`
- Create: `src/mcpOffice/Services/Excel/Vba/VbaModuleEntry.cs`

The spike's `ParseDirStream` (lines 150-200) is the working reference. Promote with two enhancements:

1. **Prefer Unicode siblings** — track both MBCS and Unicode forms; if Unicode present, use it.
2. **Handle PROJECTVERSION quirk explicitly** (already in spike — keep the comment).

```csharp
// src/mcpOffice/Services/Excel/Vba/VbaModuleEntry.cs
namespace McpOffice.Services.Excel.Vba;

internal sealed record VbaModuleEntry(
    string Name,
    string StreamName,
    uint TextOffset,
    ushort Type); // 0x0021 procedural, 0x0022 class/document
```

```csharp
// src/mcpOffice/Services/Excel/Vba/VbaDirStreamParser.cs
using System.Text;

namespace McpOffice.Services.Excel.Vba;

internal static class VbaDirStreamParser
{
    // MS-OVBA 2.3.4.2 — dir stream: sequence of (id u16 | size u32 | payload[size])
    // records. Module run starts at MODULENAME (0x0019) or MODULENAMEUNICODE (0x0047)
    // and ends at Terminator (0x002B, size 0). Other 0x002B terminators (project-section
    // ends) appear too — only emit when a module-name record was seen in the run.
    public static IReadOnlyList<VbaModuleEntry> Parse(byte[] decompressedDirStream)
    {
        // Implementation extracted from spike, plus:
        //   case 0x0047: nameUnicode = Encoding.Unicode.GetString(...)
        //   case 0x0032: streamNameUnicode = Encoding.Unicode.GetString(...)
        // On terminator: prefer Unicode where present, else MBCS.
    }
}
```

**Commit:**
```bash
git commit -am "feat: VbaDirStreamParser with MBCS+Unicode module-name handling"
```

---

### Task 6: Unit-test the `PROJECTVERSION` quirk explicitly

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaDirStreamParserTests.cs`

```csharp
[Fact]
public void Walks_past_PROJECTVERSION_record_with_real_payload_size_6()
{
    // Hand-craft a dir stream containing PROJECTVERSION (id=0x0009, size=4, payload=6),
    // followed by a complete module run. Assert the parser does NOT skip 2 bytes too few
    // (which would corrupt the next record's id/size fields).

    var stream = BuildDirStream(records:
    [
        Record(0x0009, sizeField: 4, payload: new byte[] { 0x01, 0x00, 0x00, 0x00, 0x02, 0x00 }), // PROJECTVERSION (size says 4, payload is 6)
        Record(0x0019, sizeField: 7, payload: AsMbcs("Module1")),                                 // MODULENAME
        Record(0x001A, sizeField: 7, payload: AsMbcs("Module1")),                                 // MODULESTREAMNAME
        Record(0x0031, sizeField: 4, payload: BitConverter.GetBytes(0u)),                         // MODULEOFFSET=0
        Record(0x0021, sizeField: 0, payload: []),                                                // MODULETYPE procedural
        Record(0x002B, sizeField: 0, payload: []),                                                // Terminator
    ]);

    var modules = VbaDirStreamParser.Parse(stream);

    Assert.Single(modules);
    Assert.Equal("Module1", modules[0].Name);
    Assert.Equal((ushort)0x0021, modules[0].Type);
}
```

> Helper builders (`BuildDirStream`, `Record`, `AsMbcs`) live in the test file. Keep them simple — byte-array concat. cp1252 lookup needs `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);` once per AppDomain — wrap in a `[ModuleInitializer]` or a static ctor on the test class.

Also add tests for:

- **Unicode sibling preference** — feed both MBCS (`0x0019` "ModABC") and Unicode (`0x0047` "ΜοδΑΒΓ") name records in the same module run; assert `Name == "ΜοδΑΒΓ"`.
- **Multiple modules in one stream** — two complete module runs, assert both come back in order.
- **Terminator without preceding MODULENAME does not emit** — covers the project-section terminator case.

**Commit:**
```bash
git commit -am "test: VbaDirStreamParser — PROJECTVERSION quirk, Unicode preference, multi-module"
```

---

# Phase 3 — VBA project reader (orchestrator)

### Task 7: Add VBA DTOs

**Files:**
- Create: `src/mcpOffice/Models/ExcelVbaModule.cs`
- Create: `src/mcpOffice/Models/ExcelVbaProject.cs`

```csharp
// src/mcpOffice/Models/ExcelVbaModule.cs
namespace McpOffice.Models;

public sealed record ExcelVbaModule(
    string Name,
    string Kind,       // "standardModule" | "classModule" | "documentModule"
    int LineCount,
    string Code);
```

```csharp
// src/mcpOffice/Models/ExcelVbaProject.cs
namespace McpOffice.Models;

public sealed record ExcelVbaProject(
    bool HasVbaProject,
    IReadOnlyList<ExcelVbaModule> Modules);
```

**Commit:**
```bash
git commit -am "feat: ExcelVbaProject and ExcelVbaModule DTOs"
```

---

### Task 8: Implement `VbaProjectReader` with split API

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaProjectReader.cs`

The reader exposes two entry points so the OLE/decompression core is testable without a `.xlsm` fixture:

- `Read(string xlsmPath)` — opens the zip, finds `xl/vbaProject.bin`, hands the bytes to `ReadVbaProjectBin`. Returns `ExcelVbaProject(HasVbaProject: false, ...)` when the entry is absent.
- `ReadVbaProjectBin(Stream stream, string sourceLabel)` — does the OLE walk + dir decompression + module decompression. `sourceLabel` is used only in error messages (the `.xlsm` path for production calls; a synthetic name like `"<synthetic>"` for unit tests).

Responsibilities of `ReadVbaProjectBin`:

1. `OpenMcdf.RootStorage.Open(stream)` to read the OLE compound file.
2. Open `VBA` storage, read `dir` stream, decompress via `MsOvbaDecompressor`, parse via `VbaDirStreamParser`.
3. For each module entry: open `<streamName>` under `VBA` storage, read all bytes, slice from `TextOffset`, decompress, decode as cp1252.
4. Map MODULETYPE: `0x0021` → `"standardModule"`. `0x0022` → `"documentModule"` if module name is `ThisWorkbook` or starts with `Sheet` (cheap heuristic — see footnote on `VB_Base` refinement); otherwise `"classModule"`.
5. Wrap `IOException`/`InvalidDataException`/OpenMcdf exceptions in `vba_parse_error` with the original message preserved.
6. **Locked detection:** if `dir` stream cannot be located, OR decompressing `dir` yields garbage that fails to parse as any record run → throw `vba_project_locked`. Conservative v1: detect via "dir decompresses but no module runs found AND parser produced zero records past the project-section terminator". Refine when a real locked sample arrives (Open Question #1).

Responsibilities of `Read(string)`:

1. Open `.xlsm` as ZIP via `System.IO.Compression.ZipFile.OpenRead`.
2. If no `xl/vbaProject.bin` entry → return `ExcelVbaProject(HasVbaProject: false, Modules: [])`. **Not** an error: design doc shows `hasVbaProject: false` as a normal response. The `vba_project_missing` code stays defined but unused by this tool — reserved for a future strict variant (`excel_extract_vba_required`).
3. Stream entry into a `MemoryStream` (1.17 MB peak observed; small enough for memory).
4. Delegate to `ReadVbaProjectBin(memoryStream, xlsmPath)`.

```csharp
// src/mcpOffice/Services/Excel/Vba/VbaProjectReader.cs
using System.IO.Compression;
using System.Text;
using McpOffice.Models;
using ModelContextProtocol;
using OpenMcdf;

namespace McpOffice.Services.Excel.Vba;

internal sealed class VbaProjectReader
{
    private const string VbaProjectEntryName = "xl/vbaProject.bin";

    static VbaProjectReader()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public ExcelVbaProject Read(string xlsmPath)
    {
        try
        {
            using var zip = ZipFile.OpenRead(xlsmPath);
            var entry = zip.GetEntry(VbaProjectEntryName);
            if (entry is null)
            {
                return new ExcelVbaProject(HasVbaProject: false, Modules: []);
            }

            using var ms = new MemoryStream();
            using (var s = entry.Open()) s.CopyTo(ms);
            ms.Position = 0;
            return ReadVbaProjectBin(ms, xlsmPath);
        }
        catch (McpException) { throw; }
        catch (Exception ex)
        {
            throw ToolError.VbaParseError(xlsmPath, ex.Message);
        }
    }

    public ExcelVbaProject ReadVbaProjectBin(Stream vbaProjectBin, string sourceLabel)
    {
        try
        {
            using var root = RootStorage.Open(vbaProjectBin);
            // ... walk VBA storage, decompress dir, build module list
            // Throw ToolError.VbaProjectLocked / VbaParseError as appropriate.
        }
        catch (McpException) { throw; }
        catch (Exception ex)
        {
            throw ToolError.VbaParseError(sourceLabel, ex.Message);
        }
    }
}
```

> `ReadVbaProjectBin` is `public` (not `internal`) on this `internal` class so the test project can call it directly. If we ever need to expose the reader publicly, narrow this to `internal` and add an `[InternalsVisibleTo]` attribute for the test assembly.

> Use `RootStorage.Open(Stream)` (not `RootStorage.OpenRead(string)` as in the spike) so we don't re-extract `vbaProject.bin` to disk. OpenMcdf 3.1.3 supports stream-based opens. **Verify during implementation**; if stream-based open isn't supported, fall back to writing to a temp file in `Read(string)` only — `ReadVbaProjectBin` callers can still pass any stream.

**Commit (after Task 11 tests are passing):**
```bash
git commit -am "feat: VbaProjectReader with split Read(path)/ReadVbaProjectBin(stream) API"
```

---

### Task 9: Add `VbaProjectBinBuilder` test helper

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaProjectBinBuilder.cs`

A test-only builder that constructs a synthetic `vbaProject.bin` (OLE compound file) in memory from a list of `ModuleSpec` records. Used by `VbaProjectReaderTests` to drive `ReadVbaProjectBin(Stream, ...)` without needing an `.xlsm` fixture or real Excel output.

**API:**

```csharp
internal sealed record ModuleSpec(
    string Name,
    string StreamName,
    string Source,
    bool IsDocumentModule = false,   // emit MODULETYPE 0x0022 instead of 0x0021
    string? UnicodeName = null);     // optional MODULENAMEUNICODE record

internal static class VbaProjectBinBuilder
{
    public static byte[] Build(IReadOnlyList<ModuleSpec> modules);

    // Used by tests that want to corrupt specific bytes
    public static byte[] BuildWithDirCorruption(IReadOnlyList<ModuleSpec> modules);
}
```

**Implementation outline (~120 LOC total):**

1. **Compress (literal-only):** `byte[] CompressLiteralsOnly(byte[] data)`. Emit `0x01` signature, then compressed-mode chunks where every flag-byte bit is `0` (all literals). Each chunk holds up to 4096 source bytes; chunkSize header = `(2 + (segmentLen+7)/8 + segmentLen)`. ~25 LOC. **No copy tokens needed** — the production decompressor handles literal-only compressed chunks identically to copy-token chunks.

2. **Module stream payload:** the bytes that go in `VBA/<streamName>` are:
   - Optional `Attribute VB_Name = "<name>"\r\n` (0..n bytes of attribute lines, becomes part of the compressed source body — this is the "PerformanceCache" / textOffset gap. For tests we set `textOffset = 0` and put the source itself starting at byte 0; the production reader doesn't care about the attribute lines, only `TextOffset`).
   - Compressed source via `CompressLiteralsOnly(Encoding.GetEncoding(1252).GetBytes(spec.Source))`.

3. **Dir stream payload:** sequence of records. Minimum required for the parser:
   - PROJECTSYSKIND (0x0001, size 4, payload `0x01 0x00 0x00 0x00`) — emit so we exercise the parser past the per-spec quirky records too.
   - PROJECTVERSION (0x0009, sizeField `4`, payload **6 bytes**) — exercises the spike's special-case in production.
   - For each module:
     - MODULENAME (0x0019, MBCS) and optionally MODULENAMEUNICODE (0x0047, UTF-16 LE)
     - MODULESTREAMNAME (0x001A, MBCS) and MODULESTREAMNAMEUNICODE (0x0032, UTF-16 LE)
     - MODULEOFFSET (0x0031, size 4, `textOffset = 0`)
     - MODULETYPE: `0x0022` if `IsDocumentModule`, else `0x0021` (size 0)
     - Terminator (0x002B, size 0)
   - Top-level Terminator (0x002B, size 0).

   Each record is `id (UInt16 LE) | size (UInt32 LE) | payload[size]`. ~40 LOC.

4. **OLE compound file:** use `OpenMcdf.RootStorage.Create(stream, ...)` (or equivalent OpenMcdf 3.1.3 write API) to build:
   - `VBA` storage containing:
     - `dir` stream = `CompressLiteralsOnly(dirRecords)`
     - one stream per module = `CompressLiteralsOnly(moduleSourceBytes)`
   - We can omit `VBA/_VBA_PROJECT`, `PROJECT`, `PROJECTwm` — the production reader doesn't read them. If OpenMcdf complains about an empty root, add a 1-byte `PROJECT` stream with arbitrary content. ~40 LOC.

5. Return the compound-file bytes via a `MemoryStream`.

> If OpenMcdf 3.1.3's write API turns out to require more storage entries than the bare minimum, document the actual requirement when implementing — the LOC estimate may grow by ~20. Fallback: hand-write the OLE compound-file binary format. Not recommended; OpenMcdf write should work.

**Verification:** before relying on the builder, write a smoke test that:
1. Builds a tiny project with one module.
2. Pipes the bytes through `MsOvbaDecompressor` directly on the dir stream.
3. Asserts the resulting record stream parses via `VbaDirStreamParser.Parse(...)` to the expected single module.

This catches builder bugs before they confuse reader-test failures.

**Commit:**
```bash
git commit -am "test: VbaProjectBinBuilder for synthetic vbaProject.bin construction"
```

---

### Task 10: Unit tests for `VbaProjectReader.ReadVbaProjectBin` (synthetic)

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaProjectReaderTests.cs`

These tests drive `ReadVbaProjectBin(Stream, sourceLabel)` directly with builder output. No fixture file. Fast, parameterizable, no Excel dependency.

```csharp
public class VbaProjectReaderTests
{
    [Fact]
    public void Reads_single_standard_module()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("Module1", "Module1", "Sub Hello()\r\n  Debug.Print \"hi\"\r\nEnd Sub")
        ]);

        var project = new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(bytes), "<synthetic>");

        Assert.True(project.HasVbaProject);
        var m = Assert.Single(project.Modules);
        Assert.Equal("Module1", m.Name);
        Assert.Equal("standardModule", m.Kind);
        Assert.Contains("Sub Hello", m.Code);
        Assert.Equal(3, m.LineCount);
    }

    [Fact]
    public void Reads_document_module_with_correct_kind()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("ThisWorkbook", "ThisWorkbook",
                           "Private Sub Workbook_Open()\r\nEnd Sub",
                           IsDocumentModule: true)
        ]);

        var project = new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(bytes), "<synthetic>");
        Assert.Equal("documentModule", project.Modules.Single().Kind);
    }

    [Fact]
    public void Reads_multiple_modules_in_order()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("ModA", "ModA", "' a"),
            new ModuleSpec("ModB", "ModB", "' b"),
            new ModuleSpec("ThisWorkbook", "ThisWorkbook", "' wb", IsDocumentModule: true)
        ]);

        var project = new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(bytes), "<synthetic>");
        Assert.Equal(["ModA", "ModB", "ThisWorkbook"], project.Modules.Select(m => m.Name).ToArray());
    }

    [Fact]
    public void Prefers_unicode_module_name_when_present()
    {
        var bytes = VbaProjectBinBuilder.Build([
            new ModuleSpec("ModABC", "ModABC", "' x", UnicodeName: "ΜοδΑΒΓ")
        ]);

        var project = new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(bytes), "<synthetic>");
        Assert.Equal("ΜοδΑΒΓ", project.Modules.Single().Name);
    }

    [Fact]
    public void Throws_vba_parse_error_for_corrupt_input()
    {
        var corrupt = new byte[] { 0x00, 0x01, 0x02, 0x03, 0x04 }; // not an OLE file
        var ex = Assert.Throws<McpException>(() =>
            new VbaProjectReader().ReadVbaProjectBin(new MemoryStream(corrupt), "<synthetic>"));
        Assert.Contains("vba_parse_error", ex.Message);
        Assert.Contains("<synthetic>", ex.Message);
    }

    [Fact(Skip = "needs locked-project sample — see SESSION_HANDOFF.md Open Question #1")]
    public void Throws_vba_project_locked_for_protected_project() { }
}
```

**Commit:**
```bash
git commit -am "test: VbaProjectReader.ReadVbaProjectBin against synthetic builder"
```

---

### Task 11: Hand-authored fixture + real-Excel smoke test

This task validates the zip-extraction path of `Read(string)` and the MS-OVBA compressed-chunk decoder against actual Excel-produced bytes. It's the only place an `.xlsm` fixture is required.

**Files:**
- Create: `tests/fixtures/sample-with-macros.xlsm`
- Create or extend: `tests/fixtures/README.md` documenting how the fixture was authored
- Modify: `tests/mcpOffice.Tests/Excel/Vba/VbaProjectReaderTests.cs` — add real-Excel smoke + xlsx-without-macros tests

**Authoring the fixture (manual, one-time, ~5 minutes):**

1. Open Excel → New blank workbook.
2. Save As → `sample-with-macros.xlsm` (keep it tiny — single sheet, no styles).
3. Alt+F11 → VBA editor:
   - Insert → Module → name it `Module1`. Body: `Sub Hello()`<br>`  Debug.Print "hi"`<br>`End Sub`
   - Project Explorer → double-click `ThisWorkbook`. Body: `Private Sub Workbook_Open()`<br>`End Sub`
4. Save and close. Target size: under 30 KB. Move to `tests/fixtures/sample-with-macros.xlsm`.

Add or extend `tests/fixtures/README.md`:

```markdown
## sample-with-macros.xlsm

Tiny `.xlsm` for VBA extraction tests. Contains:
- `Module1` (standard module): a `Sub Hello()` that calls `Debug.Print`.
- `ThisWorkbook` (document module): an empty `Private Sub Workbook_Open()` event handler.

Authored manually in Excel — DevExpress can't write VBA. To regenerate, follow
the steps in `docs/plans/2026-05-01-mcpoffice-excel-vba-extraction-plan.md`
Task 11.
```

**Tests to add to `VbaProjectReaderTests`:**

```csharp
[Fact]
public void Reads_modules_from_real_excel_fixture()
{
    // End-to-end: zip → vbaProject.bin → real MS-OVBA compressed chunks → modules.
    // The synthetic builder uses literal-only compressed chunks; this test ensures
    // we still parse Excel's actual (copy-token) compressed output.
    var path = TestFixtures.Path("sample-with-macros.xlsm");
    var project = new VbaProjectReader().Read(path);

    Assert.True(project.HasVbaProject);
    Assert.Contains(project.Modules, m => m.Name == "Module1" && m.Kind == "standardModule");
    Assert.Contains(project.Modules, m => m.Name == "ThisWorkbook" && m.Kind == "documentModule");
    Assert.Contains("Sub Hello", project.Modules.Single(m => m.Name == "Module1").Code);
}

[Fact]
public void Returns_HasVbaProject_false_for_xlsx_without_macros()
{
    var path = TestExcelWorkbooks.Create(wb => wb.Worksheets[0].Cells["A1"].Value = "x");
    try
    {
        var project = new VbaProjectReader().Read(path);
        Assert.False(project.HasVbaProject);
        Assert.Empty(project.Modules);
    }
    finally { File.Delete(path); }
}
```

> `TestFixtures` helper exists for Word tests at `tests/mcpOffice.Tests/TestFixtures.cs`. Reuse it.

**Why "Reads_modules_from_real_excel_fixture" matters:** the synthetic builder emits literal-only compressed chunks (top bit set, all flag bits zero). Excel emits real copy-token compressed chunks. The decompressor handles both, but only this test exercises the copy-token path against real-world output. Without it we'd never validate that production input actually works — exactly the failure mode the spike was designed to catch.

**Commit:**
```bash
git add tests/fixtures/sample-with-macros.xlsm tests/fixtures/README.md
git commit -m "test: hand-authored .xlsm fixture + real-Excel smoke test for VbaProjectReader.Read"
```

---

# Phase 4 — Service + tool wiring

### Task 12: Extend `IExcelWorkbookService` with `ExtractVba`

**Files:**
- Modify: `src/mcpOffice/Services/Excel/IExcelWorkbookService.cs`
- Modify: `src/mcpOffice/Services/Excel/ExcelWorkbookService.cs`

```csharp
// IExcelWorkbookService.cs
ExcelVbaProject ExtractVba(string path);
```

```csharp
// ExcelWorkbookService.cs
public ExcelVbaProject ExtractVba(string path)
{
    PathGuard.RequireExists(path);
    return new VbaProjectReader().Read(path);
}
```

> `PathGuard.RequireExists` covers `file_not_found` / `invalid_path`. `VbaProjectReader.Read` raises `vba_parse_error` / `vba_project_locked` and is responsible for the rest. No try/catch wrapping at the service layer — the reader already wraps.

**Test:** unit test on the service layer is redundant with `VbaProjectReaderTests`. Add **one** test that asserts file_not_found surfaces:

```csharp
[Fact]
public void ExtractVba_throws_file_not_found_for_missing_workbook()
{
    var ex = Assert.Throws<McpException>(() =>
        new ExcelWorkbookService().ExtractVba(@"C:\does-not-exist.xlsm"));
    Assert.Contains("file_not_found", ex.Message);
}
```

**Commit:**
```bash
git commit -am "feat: IExcelWorkbookService.ExtractVba"
```

---

### Task 13: Add `excel_extract_vba` MCP tool

**Files:**
- Modify: `src/mcpOffice/Tools/ExcelTools.cs`

```csharp
[McpServerTool(Name = "excel_extract_vba")]
[Description("Statically extracts VBA module source from an .xlsm workbook without launching Excel. Returns hasVbaProject and a list of {name, kind, lineCount, code}. For .xlsx or workbooks without macros, returns hasVbaProject=false and an empty list.")]
public static object ExcelExtractVba(
    [Description("Absolute path to the .xlsm workbook")] string path)
    => Service.ExtractVba(path);
```

**Commit:**
```bash
git commit -am "feat: excel_extract_vba MCP tool"
```

---

# Phase 5 — Integration

### Task 14: Update `ToolSurfaceTests`

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs`

Add `"excel_extract_vba"` to the `expected` array (alphabetical position: after `excel_read_sheet`, before `Ping`).

**Commit:**
```bash
git commit -am "test: include excel_extract_vba in tool surface assertion"
```

---

### Task 15: Stdio integration test for `excel_extract_vba`

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs`

```csharp
[Fact]
public async Task Extract_vba_via_stdio_returns_modules()
{
    var fixture = ResolveFixturePath("sample-with-macros.xlsm");

    await using var harness = await ServerHarness.StartAsync();
    var result = await harness.Client.CallToolAsync(
        "excel_extract_vba",
        new Dictionary<string, object?> { ["path"] = fixture });

    var text = result.Content.OfType<TextContentBlock>().Single().Text;

    Assert.Contains("\"hasVbaProject\":true", text);
    Assert.Contains("\"name\":\"Module1\"", text);
    Assert.Contains("\"kind\":\"standardModule\"", text);
    Assert.Contains("Sub Hello", text);
}

[Fact]
public async Task Extract_vba_via_stdio_returns_empty_for_xlsx()
{
    var path = TempPath(".xlsx");
    try
    {
        using (var workbook = new Workbook())
        {
            workbook.Worksheets[0].Cells["A1"].Value = "x";
            workbook.SaveDocument(path, SpreadsheetFormat.Xlsx);
        }

        await using var harness = await ServerHarness.StartAsync();
        var result = await harness.Client.CallToolAsync(
            "excel_extract_vba",
            new Dictionary<string, object?> { ["path"] = path });

        var text = result.Content.OfType<TextContentBlock>().Single().Text;
        Assert.Contains("\"hasVbaProject\":false", text);
    }
    finally { if (File.Exists(path)) File.Delete(path); }
}

private static string ResolveFixturePath(string name)
{
    var asmDir = Path.GetDirectoryName(typeof(ExcelWorkflowTests).Assembly.Location)!;
    var dir = new DirectoryInfo(asmDir);
    while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "mcpOffice.sln")))
        dir = dir.Parent;
    return Path.Combine(dir!.FullName, "tests", "fixtures", name);
}
```

**Commit:**
```bash
git commit -am "test: stdio integration test for excel_extract_vba (with and without macros)"
```

---

# Phase 6 — Verification + handoff

### Task 16: Final verification

**Steps:**

1. `dotnet build --nologo` — 0 warnings, 0 errors.
2. `dotnet test --nologo` — all unit + integration tests pass. Expected count: previous `47/47` + roughly `19` new (3 error-code, 5 decompressor, 4 dir-parser, 1 builder smoke, 5 reader-synthetic, 2 reader-real-excel, 1 service, 0 tool unit, 2 stdio integration; locked test skipped) ≈ `66/67 passing, 1 skipped`.
3. **Live agent verification** — wire the rebuilt server into Claude Code (existing `claude_desktop_config.json`) and call `excel_extract_vba` against `C:\temp\macro\Air - Labware.xlsm` with a real agent. Per global CLAUDE.md: build green ≠ it works; verify with a real agent call.
4. Update `SESSION_HANDOFF.md` — mark plan items 6, 7 done; note open items (locked-project fixture, `excel_analyze_vba` future work). Bump tool catalog count to 19.

**Commit:**
```bash
git commit -am "docs: handoff after excel_extract_vba lands"
```

### Task 17: Open PR back to `main`

Squash-merge per CLAUDE.md. PR title: `feat: excel_extract_vba — static VBA source extraction`.

---

## Risks called out

1. **Synthetic builder may diverge from real Excel output.** The Task 9 builder emits literal-only compressed chunks; Excel emits copy-token compressed chunks. The decompressor handles both, so the unit-test surface is sound *for what it tests*. The Task 11 real-Excel smoke test exists specifically to close this gap. If that one test passes locally but the synthetic tests don't reflect the actual structural variability of Excel-produced files, we'll see it as bug reports against real `.xlsm` inputs in the wild — at which point we either author more fixtures or extend the builder. Mitigation today: keep the spike output (`C:\temp\macro\vba-spike-output.txt`) as a manual cross-check during implementation.

2. **`vba_project_locked` detection is heuristic-only without a real locked sample.** Plan ships a conservative detector (no module runs found + protection record present). When a locked fixture arrives (Open Question #1), tighten the detector and unskip the test. Worst current behavior: a locked project surfaces as `vba_parse_error` instead of `vba_project_locked` — wrong code but still a hard fail, agent will not silently ignore.

3. **MS-OVBA spec edge cases beyond `PROJECTVERSION`.** Spike covered the main one. Other quirks (REFERENCEPROJECT records, MODULEPRIVATE, host-extensible records) may surface on workbooks with unusual references. Wrap the dir parser in defensive bounds checks (already present in spike) so unknown record types are skipped, not throw.

4. **`OpenMcdf.RootStorage.Open(Stream)` API surface** — spike uses `OpenRead(string)`. Verify stream-based open works in 3.1.3 during Task 8; if not, fall back to writing `vbaProject.bin` to a temp file and using `OpenRead(string)`. Either is fine; stream is just cleaner.

5. **Source decoding assumes cp1252.** A workbook authored on a non-Western locale may have a different MBCS code page. MS-OVBA stores the project's `LCID` in the dir stream (record `0x0002 PROJECTLCID`); the production reader can read it and pick the matching code page. **Stretch goal:** if implementing is cheap, do it; otherwise document and defer. Initial implementation: cp1252 always.

## What this plan deliberately does NOT do

- `excel_analyze_vba` (procedure/event/dependency analysis on top of extracted source). Separate future plan.
- Form-layout streams (`f`, `o`, `VBFrame` per UserForm). Source code only — design doc is explicit.
- VBA project password recovery. Out of scope, ethically and legally fraught.
- Re-exporting `MsOvbaDecompressor` / `VbaDirStreamParser` as public types. Both stay `internal`.
- Promoting the spike file. Leave `tests/mcpOffice.Tests/Spikes/VbaExtractionSpike.cs` in place as historical reference (it's tagged in the handoff). It still runs as a no-op when `C:\temp\macro\vbaProject.bin` is absent, so it won't fail in CI.
