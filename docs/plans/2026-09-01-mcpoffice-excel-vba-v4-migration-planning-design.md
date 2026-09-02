# mcpOffice — VBA analyzer v4: migration-planning tools

**Date:** 2026-09-01
**Status:** Approved (user picked items 1-4 now, 5 when time permits)
**Cards:** VBA-012 (entry points / dead code), VBA-013 (sheet access map), VBA-014 (corpus dedup), VBA-015 (UserForm controls)
**Depends on:** v1 analyzer (`excel_analyze_vba`), v2 renderer, v3 hints — all untouched.

## Purpose

v1-v3 describe the code. A migration agent planning an Excel-to-C# rewrite needs four things the
code alone does not say: what actually *runs* (so dead code is not ported), which sheet cells the
code treats as its *database*, which code is *shared* across the lab's workbooks (so it is ported
once), and what the *UserForms* look like. v4 adds four tools that compose the existing pipeline
plus the Excel read services. No new parsing strategy: regex on cleaned source (v1), Open XML
walks (VBA extractor), DevExpress for formulas and defined names (`excel_list_formulas`,
`excel_list_defined_names`).

Design rule (ARCHITECTURE.md, "Design for an LLM caller"): each tool is opt-in and scoped, every
record self-contained, vocabularies closed and written here, heavy arrays capped.

## Why separate tools and not toggles on `excel_analyze_vba`

Each needs inputs v1 does not load (sheet drawing parts, formulas, other workbooks) and has its
own bounded output. v1/v2/v3 stay unopinionated and untouched, as the v3 design promised. Tool
count goes 34 → 38 (37 without VBA-015).

## What DevExpress does and does not give us (dxdocs 26.1, verified 2026-09-01)

- `Worksheet.FormControls` / `Worksheet.Shapes` expose form controls and shapes with `Name`,
  `FormControlType`, position — but **no macro link**. `ButtonFormControl` has `PlainText`,
  `PrintObject`, `Hyperlink`, nothing like `OnAction`.
- Formulas and defined names could come from the DevExpress-backed `excel_list_formulas` /
  `excel_list_defined_names`, but that loads the whole workbook (30 s on ScreeningDB-V2). v4 reads
  both straight from `xl/worksheets/sheetN.xml` and `xl/workbook.xml` instead (`OpenXmlParts`),
  so none of the three tools touches DevExpress at all. *(Implementation note, 2026-09-02.)*
- Macro wiring comes from the Open XML drawing parts (below). Shape names come along for free
  from the same parts, so v4 does not use the DevExpress shape API either.

## Tool 1 — `excel_list_vba_entry_points` (VBA-012)

`excel_list_vba_entry_points(path, includeUnreachable=true, moduleName?)`

### Entry-point sources (closed set for `kind`)

| kind | Where it comes from |
|---|---|
| `eventHandler` | v1 `IsEventHandler` (sheet/workbook/form/ActiveX code-behind) |
| `autoMacro` | `Auto_Open`, `Auto_Close`, `Auto_Activate`, `Auto_Deactivate` in standard modules |
| `shapeMacro` | `xl/drawings/drawingN.xml`: `macro="…"` on `<xdr:sp>`, `<xdr:pic>`, `<xdr:cxnSp>`, and children of `<xdr:grpSp>` |
| `formControlMacro` | `xl/drawings/vmlDrawingN.vml`: `<x:ClientData ObjectType="Button|Checkbox|Drop|Radio|Spin|Scroll|List|Label|GBox"><x:FmlaMacro>` |
| `worksheetFunction` | `Public Function` in a standard module whose name appears as `Name(` in any cell formula (formulas streamed from the sheet XML via `OpenXmlParts.ReadFormulas` — not `excel_list_formulas`, which would load the workbook through DevExpress: 30 s on ScreeningDB-V2 for nothing; case-insensitive, not preceded by `.` or a word char) |
| `dynamicDispatch` | string-literal targets of `Application.OnTime`, `Application.OnKey`, `Application.Run`, `.OnAction = "…"` inside VBA — recorded as an edge from the containing procedure *and* listed here |

Macro references in drawing parts have the form `[0]!Module.Proc`, `Module.Proc`, `Proc`, or
`'Book.xlsm'!Proc`. Resolution: strip the `[n]!` / `'…'!` prefix; `Module.Proc` resolves to that
FQN; bare `Proc` resolves to the unique procedure with that name (case-insensitive) across
standard modules; ambiguous or missing → `resolved: false`, `target` kept verbatim.

Sheet attribution: `xl/workbook.xml` `<sheet name r:id>` → `xl/_rels/workbook.xml.rels` →
`xl/worksheets/sheetN.xml` → `<drawing r:id>` / `<legacyDrawing r:id>` →
`xl/worksheets/_rels/sheetN.xml.rels` → drawing part. Standard `System.IO.Compression` +
`System.Xml.Linq`, like `VbaProjectReader`'s vbaProject.bin lookup.

### Reachability

BFS from every resolved entry point over v1's FQN call graph plus the `dynamicDispatch` edges.
`unreachable[]` = procedures never visited. Confidence per record:

- `high` — `Private` procedure in a standard module, no `CallByName` / non-literal `Application.Run`
  anywhere in the workbook.
- `medium` — everything else (class-module members reached through object variables, `Public`
  procedures that an external workbook could call, or any dynamic dispatch with a non-literal
  target present — then `summary.dynamicDispatchUnresolved > 0` says why).

Property procedures (`Get/Let/Set`) are treated like methods. Document-module procedures that are
not event handlers are ordinary candidates.

### Output

```
{
  "path", "hasVbaProject",
  "summary": { "entryPointCount", "byKind": {kind: n}, "procedureCount", "reachableCount",
               "unreachableCount", "unresolvedMacroReferences", "dynamicDispatchUnresolved" },
  "entryPoints": [ { "procedure": "Module.Proc", "kind", "sheet"?, "shapeName"?, "cell"?,
                     "formulaCells"?: ["Sheet!A1", …] (first 5), "target"?, "resolved": true } ],
  "unreachable":  [ { "procedure", "module", "moduleKind", "scope", "lineCount", "confidence" } ],
  "truncated": false
}
```

Caps: `entryPoints` and `unreachable` sorted (module, procedure); `maxItems` 500 each,
`truncated: true` when cut; `moduleName` scopes both arrays (summary stays whole-workbook, as v1).

### Errors

Existing codes only: `file_not_found`, `invalid_path`, `parse_error`, `vba_project_locked`.
Workbooks without a VBA project return `hasVbaProject: false` with a zeroed summary (v1 convention).
Drawing parts that fail to parse are skipped and counted in `summary.skippedDrawingParts`; never fatal.

## Tool 2 — `excel_map_vba_sheet_access` (VBA-013)

`excel_map_vba_sheet_access(path, moduleName?, sheetName?, includeUnresolved=true, includeRecords=true, maxRecords=300)`

### Resolution rules (regex on cleaned lines, `With`-aware)

| Pattern | Resolves to |
|---|---|
| `Worksheets("X")`, `Sheets("X")`, `ThisWorkbook.Worksheets("X")` | sheet by name |
| `Blad1.…`, `Sheet1.…` (document-module codename) | sheet by codename → name via `<sheetPr codeName>` in `xl/worksheets/sheetN.xml` |
| `.Range("A1:B2")`, `.Cells(r, c)`, `[A1]`, `.Columns("A:B")`, `.Rows(5)` | range on the qualifying sheet; `Cells(r, c)` with non-literal args → target `dynamicCells` |
| `Range("Name")`, `Names("Name")`, `[Name]` where `Name` is a defined name | `excel_list_defined_names` → `refersTo` gives the real sheet + range; `target.definedName` kept |
| `ActiveSheet.…`, unqualified `Range(`/`Cells(` outside a `With` | sheet `null`, `unresolvedReason: "activeSheet"` — never guessed |
| `With <sheet expr>` … `End With` | innermost `With` qualifies leading-dot members |
| `Set ws = Worksheets("X")` then `ws.Range(...)` | one-assignment alias tracking within a procedure; reassigned aliases → `unresolvedReason: "aliasReassigned"` |

Mode: `write` when the site is an assignment target (`= ` to its right at nesting depth 0) or
the member is `.Value =`, `.Formula =`, `.Clear`, `.ClearContents`, `.Delete`, `.Insert`,
`.Copy`/`.PasteSpecial` destination, `.AutoFilter`; `read` otherwise; a procedure with both on the
same target gets `both`.

### Output

```
{
  "path", "hasVbaProject",
  "summary": { "siteCount", "resolvedCount", "unresolvedCount", "sheetCount", "procedureCount" },
  "sheetAccess": [ { "procedure": "Module.Proc", "sheet": {"name","codeName"} | null,
                     "target": {"kind": "range|definedName|column|row|wholeSheet|dynamicCells",
                                "address"?, "definedName"?},
                     "mode": "read|write|both", "siteCount", "unresolvedReason"? } ],
  "sheets": [ { "name", "codeName", "readers": ["Module.Proc"], "writers": ["Module.Proc"],
                "readSites", "writeSites" } ],
  "truncated": false
}
```

`sheetAccess` is one record per (procedure, sheet, target, mode); capped at `maxRecords` (default
300, `truncated: true` when cut), `moduleName` / `sheetName` scope it, and `includeRecords=false`
drops it altogether, leaving the summary and the rollup — the first call on a big workbook.
`sheets` is the per-sheet rollup and is never cut. *(VBA-016, 2026-09-02: the cap was a hidden
1000 and Air.xlsm returned 672 records = 114 KB on one line, more than the MCP client shows the
caller; the rollup alone is 9 KB.)*

## Tool 3 — `excel_compare_vba_corpus` (VBA-014)

`excel_compare_vba_corpus(paths?: string[], directory?: string, minOccurrences=2, maxProcedures=200, includeNearDuplicates=true)`

Exactly one of `paths` / `directory`; `directory` picks `*.xlsm` (the extractor's format),
non-recursive. Fewer than two readable workbooks → `invalid_path`.

Normalisation before hashing: `VbaLineCleaner` output (comments and blank lines gone), whitespace
collapsed to one space, whole line case-folded (VBA is case-insensitive), `Attribute` lines
dropped. Hash = SHA-256 of the joined lines. Procedure identity = normalised body; the *name* is
not part of the hash so a renamed copy still groups.

Tiers:

1. `identical` — same hash in ≥ `minOccurrences` workbooks.
2. `nearDuplicate` — same procedure name in ≥ `minOccurrences` workbooks, hashes differ,
   similarity ≥ 0.9 where similarity = 2·|common normalised lines (multiset)| / (|a| + |b|).
   `// ponytail: line-multiset similarity, not LCS; upgrade if reordered bodies matter.`

### Output

```
{
  "workbooks": [ { "path", "moduleCount", "procedureCount", "hasVbaProject", "error"? } ],
  "summary": { "workbookCount", "procedureCount", "sharedProcedureCount", "identicalGroups",
               "nearDuplicateGroups", "sharedModuleCount" },
  "sharedProcedures": [ { "tier": "identical|nearDuplicate", "name", "lineCount",
                          "occurrences": [ { "workbook", "module", "procedure", "similarity" } ] } ],
  "sharedModules": [ { "module", "workbooks": [...], "sharedProcedureRatio" } ],
  "truncated": false
}
```

Sorted by occurrence count desc, then name; `maxProcedures` caps `sharedProcedures`. A module is
"shared" when ≥ 50 % of its procedures are in some shared group across the same workbook set.
Per-workbook failures (locked project, corrupt file) land in `workbooks[].error` and the run
continues. Loading is sequential; `ScreeningDB-V2.xlsm` alone takes ~30 s, so the tool
description says "minutes for a directory".

## Tool 4 — `excel_list_vba_form_controls` (VBA-015, when time permits)

`excel_list_vba_form_controls(path, formName?)`

Inference from the `frm*` code-behind only (the binary `.frx` designer half is out of scope):

- `Me.<ctrl>`, bare `<ctrl>.<Prop>` where `<ctrl>` is not a declared variable → control name.
- `<ctrl>_<Event>(…)` handler → control + event; event → type hint: `Click` → `button?`,
  `Change`/`KeyPress` → `textBox?`/`comboBox?`, `AfterUpdate` → `textBox?`, `DblClick` on a
  `lst*` name → `listBox?`. Name prefix (`txt`, `cmd`, `btn`, `lst`, `cbo`, `chk`, `opt`, `lbl`,
  `frm`) wins over the event hint when present. The VBE default names (`Label2`, `TextBox1`,
  `CommandButton3` — MSForms type name + number) count as prefixes too; OlieGC's `Label2_Click`
  came back as a CommandButton before that (acceptance 2026-09-02).
- `Dim … As MSForms.<Type>` / `As <Type>` with `MSForms` types → exact type.

Output per form: `controls[] = { name, inferredType, typeConfidence: "declared|prefix|event|member|none",
events[], referencedProperties[] }`, `formEvents[]` (the form's own `UserForm_*` handlers — replaces the
planned `handlersWithoutControl[]`, which had no clean definition once a handler name alone implies
the control exists), `handlerCount`, `summary`. A `member` confidence (`.AddItem` → ListBox,
`.Caption` → Label, …) beats an event hint when they disagree. *(Implemented 2026-09-02.)*
Real-file check: `OlieGC - LABWARE PRD.xlsm`, `QQQ2 - Absolute.xlsm`.

## Shared plumbing added once

- `Services/Excel/OpenXmlParts.cs` — opens the package, maps sheet name ↔ `sheetN.xml` ↔ codename
  ↔ drawing / legacy-drawing parts. Used by tools 1 and 2. Pure functions over `ZipArchive`,
  unit-tested with a programmatic `.xlsm`? — DevExpress cannot author drawings with macros, so the
  drawing-part parser is tested on hand-written XML strings, and the package walk on a small
  committed fixture under `tests/fixtures/` (the one case ARCHITECTURE.md allows binaries for).
- `Services/Excel/Vba/VbaCallGraphReachability.cs` — BFS over `ExcelVbaCallEdge[]`. Pure.
- `Services/Excel/Vba/VbaSheetAccessResolver.cs` — the `With` / alias / defined-name resolver. Pure.
- `Services/Excel/Vba/VbaProcedureHasher.cs` — normalisation + hash + similarity. Pure.

## Tests

Unit (pure classes, string fixtures): drawing-part macro extraction incl. `[0]!`, `'Book'!` and
bare forms; VML `FmlaMacro`; reachability incl. cycles and dynamic edges; sheet-access resolver
per pattern row above incl. nested `With` and alias reassignment; hasher normalisation, tier
assignment, similarity threshold. Integration: `ToolSurfaceTests` gets the new names; one
`ExcelWorkflowTests` round-trip per tool. Gated real-world: Air.xlsm and RingOnderzoek.xlsm
(buttons, `Blad` codenames) for tools 1-2, the samples directory for tool 3 — skip when absent,
like `AirSampleAnalysisTests`.

## Implementation order

1. `OpenXmlParts` + drawing/VML macro extraction (pure) → tests.
2. `VbaCallGraphReachability` → tests.
3. `excel_list_vba_entry_points` service + tool + surface test + Air/RingOnderzoek gated test.
4. `VbaSheetAccessResolver` → tests, then `excel_map_vba_sheet_access` + tool.
5. `VbaProcedureHasher` → tests, then `excel_compare_vba_corpus` + tool.
6. README / usage / ARCHITECTURE domain table; roadmap item 7.
7. VBA-015 when time permits.

## Out of scope

- Parsing `.frx` (MS-OFORMS) — VBA-015 full version, separate card if the cheap one falls short.
- Cross-workbook call resolution (`'Other.xlsm'!Proc`) — recorded verbatim as unresolved.
- A VBA tokenizer. Same stance as ARCHITECTURE.md: revisit only if regex is defeated on the corpus.
- Louvain clustering (VBA-001) — `sharedModules` and v3 coupling already cluster.

## Open questions deferred to implementation

- Whether `IsEventHandler` in `VbaProcedureScanner` already recognises ActiveX control handlers
  (`CommandButton1_Click` in a sheet module); if not, tool 1 adds the `_<Event>` pattern on
  document modules as `eventHandler`.
- Exact `sheetPr codeName` availability in every sample (older files sometimes omit it; then the
  VBA `PROJECT` stream's `Document=Blad1/&H…` lines are the fallback for the codename list, and the
  name mapping is by sheet order).
