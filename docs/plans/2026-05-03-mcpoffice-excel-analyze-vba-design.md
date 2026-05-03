# mcpOffice — `excel_analyze_vba` Design

**Date:** 2026-05-03
**Status:** Approved
**Scope:** v1 of `excel_analyze_vba` — a structural analyzer that layers facts on top of `excel_extract_vba`'s raw module source. Designed for AI-agent consumers (Claude Code and similar). Conversion hints deferred to v2.

## Purpose

Help an agent answer two related questions about a macro-enabled Excel workbook:

- **Migration planning:** what would it take to rewrite this workbook's VBA in C#? Where are the entry points, what does each procedure call, what external resources does it touch, what's the blast radius of a change?
- **Understanding:** what does this thing actually do? Which sheets and ranges does it manipulate, what triggers it (button click, sheet event, manual call), what's the call graph of "if I press this, the following runs"?

Both consumers want the same structural facts. v1 surfaces those facts; v2 will add opinionated conversion hints once we have evidence from real workbooks.

## Architecture

`excel_analyze_vba` reuses the existing `VbaProjectReader` (under `src/mcpOffice/Services/Excel/Vba/`) to extract module source, then a new `VbaSourceAnalyzer` layers structural analysis on top. One MCP tool method, one synchronous call, no chaining required from the agent's side.

```
excel_analyze_vba (Tool)
  ↓
ExcelWorkbookService.AnalyzeVba(path, options)
  ↓
VbaProjectReader.Read(path)            ← existing
  ↓ ExcelVbaProject { modules[] }
VbaSourceAnalyzer.Analyze(project, options)  ← new
  ↓ ExcelVbaAnalysis
```

The analyzer is internally split into:

- `VbaLineCleaner` — turns raw source into `IReadOnlyList<CleanedLine>` with comments stripped, string literals replaced with `<STR>` sentinels (originals preserved on the side), and `_` continuations folded.
- `VbaProcedureScanner` — walks cleaned lines to identify procedure boundaries and signatures.
- `VbaCallGraphBuilder` — runs over each procedure's body to extract callee names, resolves them against the procedure index.
- `VbaReferenceCollector` — runs over each procedure's body to collect Excel object-model references and external dependencies.

Each subcomponent is independently unit-testable against synthetic source strings.

## Tool surface

```
excel_analyze_vba(
    path,                       // absolute path to .xlsm / .xlsb
    includeProcedures = true,   // per-module procedure list (cheap)
    includeCallGraph = false,   // who-calls-who edges (medium)
    includeReferences = false   // object-model + dependency refs (heavy)
)
```

Single tool. Mirrors `excel_get_structure`'s tiered-toggle pattern. `path` is absolute (rejected with `invalid_path` otherwise — same as every other tool).

The agent's expected workflow:

1. Call with all toggles default (only `includeProcedures = true`) — get a structural picture and the summary counters.
2. Decide based on counters whether to pull the call graph, references, or both.

## Output DTO

```
ExcelVbaAnalysis {
  hasVbaProject: bool
  summary: {
    moduleCount: int,
    parsedModuleCount: int,
    unparsedModuleCount: int,
    procedureCount: int,
    eventHandlerCount: int,
    callEdgeCount: int,
    objectModelReferenceCount: int,
    dependencyCount: int
  }
  modules: [                                  // present when includeProcedures = true
    {
      name: string,
      kind: string,                           // see Enum vocabularies
      parsed: bool,                           // false when the body was skipped (e.g., module_too_large)
      reason: string?,                        // populated when parsed = false
      procedures: [
        {
          name: string,
          fullyQualifiedName: string,         // "<module>.<name>", convenient join key for callGraph
          kind: string,                       // see Enum vocabularies
          scope: string?,                     // "Public" | "Private" | "Friend" | null (default Public)
          parameters: [
            { name, type?, byRef: bool, optional: bool, defaultValue? }
          ],
          returnType: string?,                // for Function and Property Get
          lineStart: int,                     // 1-based inclusive
          lineEnd: int,                       // 1-based inclusive
          isEventHandler: bool,
          eventTarget: string?                // e.g. "Workbook", "Worksheet", "CommandButton1"
        }
      ]
    }
  ]
  callGraph: [                                // present when includeCallGraph = true
    {
      from: string,                           // procedure fullyQualifiedName
      to: string,                             // procedure fullyQualifiedName, or external name
      resolved: bool,                         // true if `to` matches a known procedure in this workbook
      site: { module, procedure, line }
    }
  ]
  references: {                               // present when includeReferences = true
    objectModel: [
      { module, procedure, line, api: string, literal: string? }
    ],
    dependencies: [
      { module, procedure, line, kind: string, target: string?, operation: string? }
    ]
  }
}
```

### Enum vocabularies (closed sets)

Documented here so an agent reading one sample response can infer the schema.

- **Module `kind`:** `"standardModule"`, `"classModule"`, `"documentModule"`, `"userForm"`.
- **Procedure `kind`:** `"Sub"`, `"Function"`, `"PropertyGet"`, `"PropertyLet"`, `"PropertySet"`.
- **Procedure `scope`:** `"Public"`, `"Private"`, `"Friend"`, or `null` (defaults to `Public`).
- **Object-model `api`:** `"Worksheets"`, `"Sheets"`, `"Range"`, `"Cells"`, `"ActiveSheet"`, `"ActiveWorkbook"`, `"ThisWorkbook"`, `"Application"`, `"Selection"`, `"Names"`. Closed set for v1.
- **Dependency `kind`:** `"file"`, `"database"`, `"network"`, `"automation"`, `"shell"`. `automation` catches `CreateObject` / `GetObject` calls that aren't file/DB/network (e.g., `Outlook.Application`, `WScript.Shell`); `shell` catches the `Shell()` builtin.
- **Module `reason` (when `parsed = false`):** `"module_too_large"`, `"empty_source"`. Extensible later.

### Agent-ergonomic choices

- **`fullyQualifiedName`** on each procedure means the agent doesn't have to join `modules[].procedures[]` to look up a `callGraph` edge — the `from` / `to` strings match it directly.
- **`summary.parsedModuleCount` and `unparsedModuleCount`** let the agent see coverage at a glance, no arithmetic.
- **`resolved: bool`** on call edges separates internal calls (migrate together) from external (likely a builtin or runtime-resolved name).
- **`literal`** on object-model refs captures the first string-literal argument when present (e.g., `Worksheets("Data")` → `"Data"`). Lets the agent build a list of touched sheets without doing source surgery.

## Parsing strategy

Two-stage pipeline per module.

### Stage 1 — `VbaLineCleaner`

Walks raw source line by line and produces `CleanedLine { lineNumber, text, originalText }` where `text` has comments stripped and string literals replaced with `<STR>`.

- **`'` comments:** strip from the apostrophe to end-of-line, unless inside a string literal.
- **`Rem ` comment statement** (legacy form): same as `'`, but only when `Rem` is the leading token of a statement.
- **`"..."` strings:** replace contents with `<STR>` but keep the surrounding quotes so subsequent regex still sees a string-shaped token. Doubled `""` is the VBA escape for a literal quote — handled.
- **`_` line continuation:** when a line ends with `<whitespace>_<eol>`, concatenate the next line into one `CleanedLine`. The `lineNumber` of the merged record is the start line.
- **VBA attribute lines** (`Attribute VB_Name = ...`): pass through; the procedure scanner ignores them (they live above `Sub`/`Function`).

### Stage 2 — Regex passes over cleaned lines

- **Procedure boundaries:**
  - Open: `^\s*(Public|Private|Friend)?\s*(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+(\w+)\s*\(([^)]*)\)(\s+As\s+\w+)?`
  - Close: `^\s*End\s+(Sub|Function|Property)\s*$`
  - Lines between = procedure body; lines outside any `Sub`/`Function`/`Property` = module-level (declarations, attributes — recorded but not a `Procedure`).

- **Parameters:** split the captured parameter group on commas (respecting `<STR>` sentinels), parse each as `[ByRef|ByVal] [Optional] name [As type] [= default]`.

- **Event handlers:** procedure name matches `<Target>_<Event>` AND the containing module is `documentModule` / `classModule` / `userForm`. The `<Target>` is `Workbook`, `Worksheet`, or a control name (UserForm controls). v1 does *not* introspect the userForm layout to know which controls exist — pattern-match on the name only and let `eventTarget` carry whatever the prefix is.

- **Call graph:** within a procedure body, match
  - `^\s*(Call\s+)?(\w+)\s*(\(|$)` — direct call.
  - `Application\.Run\s+"([^"]+)"` — late-bound call by name.
  Resolve the callee against the procedure index (matching by both bare name and `module.name`); `resolved = false` if no match.

- **Object-model refs:** match cleaned lines against the closed token set (`Worksheets`, `Sheets`, `Range`, `Cells`, `ActiveSheet`, `ActiveWorkbook`, `ThisWorkbook`, `Application`, `Selection`, `Names`). Capture the `originalText`'s first string-literal argument when the API takes one.

- **Dependencies:** matched against a small dispatch table.
  - `file`: `Open`, `Kill`, `Name … As`, `MkDir`, `RmDir`, `ChDir`, `Dir`, `FileSystemObject`, `Workbooks.Open`, `Workbooks.OpenText`, `Scripting.FileSystemObject`.
  - `database`: `ADODB.Connection`, `ADODB.Recordset`, `DAO.`, `OpenDatabase`, `Workspaces(0).OpenDatabase`.
  - `network`: `MSXML2.XMLHTTP`, `WinHttp.WinHttpRequest`, `URLDownloadToFile`, `InternetExplorer.Application`.
  - `automation`: any other `CreateObject("...")` / `GetObject("...")` not matched above. Captures the ProgID literal in `target`.
  - `shell`: `Shell(`.

### Known limits (acceptable for v1)

- **Calls via `Application.Run someStringVariable`** (callee built at runtime) appear as `resolved = false, to = "<dynamic>"`.
- **`With Worksheets("X") ... .Range("Y")` blocks** — the `.Range` ref is captured but its container is lost. The agent can correlate by line proximity if needed.
- **Dynamically-built ProgIDs** (`CreateObject(progIdVar)`) appear with `target = null`.
- **Procedure declarations split across `_` continuations** — handled by the cleaner so the regex sees one line.

### Pragmatic limits

- **Per-module hard cap: 5,000 cleaned lines.** Beyond that the analyzer skips the body, records `{name, parsed: false, reason: "module_too_large"}`, and moves on.
- **No timeout.** `Air.xlsm` (107 modules, ~500 KB source) should complete well under one second on this strategy. If real-world workbooks blow that we add a budget later.

## Error model

Three tiers:

1. **Whole-call fast failure** — same set as `excel_extract_vba`: `file_not_found`, `invalid_path`, `vba_project_locked`, `vba_parse_error`. Raised by the OLE / dir / decompression layer below the analyzer; the analyzer never runs in these cases.

2. **`hasVbaProject = false`** — file opened, no VBA project found. Returns the summary with everything zero, all collections empty. Not an error.

3. **Per-module survivable failure** — analyzer hit a module it couldn't parse. The module appears in `modules` with `parsed: false, reason: "..."` and an empty `procedures` array; analysis continues on the rest. Counters in `summary` reflect what was successfully parsed; `summary.unparsedModuleCount` makes the partial-result state explicit.

The analyzer never throws on a single bad module — that would make a 107-module workbook fragile to one weird file.

## Testing

Three layers, mirroring the existing VBA extraction tests.

### Unit — synthetic source strings

Drive `VbaLineCleaner`, `VbaProcedureScanner`, `VbaCallGraphBuilder`, `VbaReferenceCollector` with hand-written VBA fragments. Each is a focused single-fact test. Coverage targets:

- Comment inside a string: `s = "This isn't a comment"`.
- String inside a comment: `' he said "hi"`.
- `_` continuation in a procedure declaration, in a parameter list, in a body line.
- `Property Get` / `Property Let` / `Property Set` parsed as separate procedure kinds.
- `Optional ByVal x As String = "default"`.
- Event handler naming: `Workbook_Open`, `Worksheet_Change`, `cmdSubmit_Click`.
- `Call Foo` vs. `Foo` vs. `Application.Run "Foo"`.
- Object-model literal capture: `Worksheets("Data")`, `Range("A1:B10")`, `Cells(1, 2)`.
- Dependency dispatch: `Open "C:\file.txt" For Input As #1`, `CreateObject("ADODB.Connection")`, `Shell("notepad.exe")`.
- Unknown `CreateObject` ProgID lands in `automation`.

### Unit — synthetic `vbaProject.bin` builder

Extend the existing `VbaProjectBinBuilder` (under `tests/mcpOffice.Tests/Excel/Vba/`) with multi-module fixtures, run the full `extract → analyze` pipeline in-process. Asserts the analyzer composes correctly with the extractor and that real OLE / decompression output flows through.

### Real-world benchmark

A `[Fact]` against `C:\Projects\mcpOffice-samples\Air.xlsm`, gated on file existence (same skip-when-absent pattern as the existing `Extract_vba_via_stdio_returns_modules` test). Earns its keep because the analyzer's surface is much wider than the extractor's; a real workbook is the only realistic regression catcher.

Asserts:

- No exceptions.
- Every module either has `parsed = true` or carries a `reason`.
- `summary.procedureCount` exceeds a plausible floor for 107 modules of real macro code.
- Call graph is non-empty.
- References include at least one `Worksheets` and at least one `Range`.

### Stdio integration

One test in `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs` — call `excel_analyze_vba` against the synthetic `tests/fixtures/sample-with-macros.xlsm`, assert `hasVbaProject = true` and that the JSON has the expected top-level keys (`summary`, `modules`). Proves the protocol layer doesn't drop anything; correctness is covered by the unit layer.

### Tool surface

Update `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs` to include `excel_analyze_vba` (today: 23 tools, will become 24).

## Out of scope for v1

- **Conversion hints** — opinionated suggestions like "this VBA Sub maps to a C# class with method X". Designed in v2 once we have evidence from real workbooks.
- **Form layout** — `userForm` modules' control hierarchy. Procedure detection still works (controls' event handlers parse fine), but the form's visual layout is not surfaced.
- **Type library resolution** — references to external libraries (e.g., a custom COM DLL) are recorded as `automation` with the ProgID, not resolved to actual types/methods.
- **Cross-workbook calls** (e.g., `Application.Run "OtherWorkbook.xlsm!Foo"`) — captured as a call edge with `resolved = false`. v1 doesn't open the target workbook to resolve.
- **Locked VBA projects** — same situation as `excel_extract_vba`: surfaces as `vba_project_locked`. Not unique to the analyzer.
- **Concurrent / async analysis** — synchronous and per-module sequential. The expected payload sizes don't justify parallelism yet.
