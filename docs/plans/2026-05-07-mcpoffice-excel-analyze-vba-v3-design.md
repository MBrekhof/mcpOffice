# mcpOffice — `excel_suggest_vba_conversion` (analyzer v3) Design

**Date:** 2026-05-07
**Status:** Approved (brainstorming phase)
**Scope:** Conversion-hints layer over `excel_analyze_vba`. Ships as a new MCP tool that consumes the v1 structural model and emits per-procedure migration hints plus workbook-wide module coupling metrics.

## Purpose

The v1 analyzer (`excel_analyze_vba`) emits the structural model of a VBA project — procedures, signatures, call graph, Excel object-model references, external dependencies. The v2 renderer (`excel_render_vba_callgraph`) gives that graph a visual surface. v3 adds the missing third pillar: **opinions** — what each procedure *means* for an Excel→C# migration, and which modules belong together.

Target consumer is dual: a code-generation agent (needs structured emission targets it can act on) and a human reviewer / planning agent (needs prose rationale to triage). v3 emits both in the same payload — the structural fields drive code-gen, the rationale gives the human narrative.

v3 deliberately stops short of *being* the code generator. It gives the downstream LLM enough scaffolding to do that work in a follow-up step.

## Operation model

Same stateless / file-path shape as v1 and v2. Internally calls `VbaSourceAnalyzer` (the v1 path), then runs a new `VbaConversionHintBuilder` over the analyzer's model. **No new parsing.** Every hint and every coupling number is derived from data v1 already collects.

## Tool surface

```
excel_suggest_vba_conversion(
    path: string,                 // absolute path to .xlsm/.xlsb
    moduleName?: string,          // case-insensitive; same semantics as excel_analyze_vba
    targetParadigm?: string       // null | "classLibrary" | "workerService" | "webApi" | "console"
) -> ConversionHints
```

Filter semantics match v1: `moduleName` filters which procedures get hints emitted. The coupling block stays whole-workbook regardless of the filter — partial coupling numbers would mislead, since you can't compute `Ca` for module X with only X's procedures in scope.

This tool joins as the 26th MCP tool. v1 and v2 stay unopinionated and untouched.

## Output schema (hints-only + light identity context)

```jsonc
{
  "summary": {
    "totalProcedures": 200,
    "hintedProcedures": 200,        // == totalProcedures unless moduleName filter
    "moduleCount": 107,
    "targetParadigm": "workerService" | null,
    "wallTimeMs": 142
  },

  "procedureHints": [
    {
      // identity (light context — enough to act on without re-joining v1 payload)
      "module": "OrderProcessing",
      "procedureName": "ProcessOrder",
      "kind": "sub" | "function" | "propertyGet" | "propertyLet" | "propertySet",
      "isEventHandler": false,
      "paramCount": 2,
      "callerCount": 5,
      "calleeCount": 3,

      "axes": {
        "trigger":  "calledOnly" | "eventHandler" | "macroEntryPoint",
        "purity":   "pure" | "readsState" | "writesState" | "sideEffectful",
        "shape":    "leaf" | "orchestrator",   // omitted when 1 <= calleeCount <= 2
        "dependencies": ["excelObjectModel", "filesystem"]   // possibly empty, always present
      },

      "rationale": "Pure scalar transform — no side effects, no Excel object model, takes (String, Double) returns Double. Leaf in call graph.",

      "csharpSuggestion": {
        "targetType": "staticMethod" | "instanceMethod" | "backgroundService" | "apiAction" | "consoleEntryPoint" | "requiresManualReview",
        "suggestedClassName": "OrderProcessor",
        "suggestedMethodName": "ProcessOrder",
        "lifetime": "static" | "scoped" | "singleton" | null,
        "isPublic": true,
        "blockers": []
      } | null
    }
  ],

  "moduleCoupling": [
    {
      "module": "OrderProcessing",
      "ca": 12,                              // afferent: incoming calls from other modules
      "ce": 4,                               // efferent: outgoing calls to other modules
      "instability": 0.25,                   // I = ce / (ca + ce); 0.0 when both 0
      "internalEdges": 18                    // intra-module edges
    }
  ],

  "couplingPairs": [
    { "from": "OrderProcessing", "to": "DbAccess", "edgeCount": 7 }
  ]
}
```

Two contract notes:

- `csharpSuggestion: null` (key present, value null) when no `targetParadigm` was passed. Keeps consumer code uniform (`proc.csharpSuggestion?.targetType`).
- `axes.shape` is omitted when the procedure has 1–2 callees — neither `leaf` nor `orchestrator` fits cleanly there, so we don't force a label.

## Classification axes — computation rules

All four axes are pure functions of v1's existing data (`Procedure`, `CallEdge`, `ObjectModelReference`, `ExternalDependencyReference`). No new VBA parsing.

### `trigger`

First match wins:
1. `eventHandler` — when `Procedure.IsEventHandler == true` (already classified by v1).
2. `macroEntryPoint` — when the procedure is **public** AND has zero callers in the call graph AND lives in a non-`documentModule` kind (i.e. plausibly invokable externally — toolbar, `Application.Run`, shortcut).
3. `calledOnly` — otherwise. Includes private orphans (dead code) — flagged honestly rather than mislabeled.

### `purity`

Derived from the procedure's object-model and external-dependency references plus a cheap regex over the procedure's source range for module-scope writes:
- `pure` — no `ObjectModelReference`, no `ExternalDependencyReference`, no module-scope writes.
- `readsState` — only read-mode `ObjectModelReference`s OR module-scope reads, no writes anywhere.
- `writesState` — at least one write-mode `ObjectModelReference` OR module-scope write, no external I/O.
- `sideEffectful` — any `ExternalDependencyReference` (filesystem / DB / network / registry / shell). Strongest label, supersedes the rest.

`Mode == "read"|"write"` already exists on `ObjectModelReference` per call site.

### `shape`

From the call graph:
- `leaf` — `calleeCount == 0`.
- `orchestrator` — `calleeCount >= 3`.
- 1 or 2 callees — `shape` axis is omitted from the response.

### `dependencies`

Sorted, deduped subset of `{excelObjectModel, filesystem, database, network, registry, shell}`:
- `excelObjectModel` if any `ObjectModelReference` exists for the procedure.
- The other five come straight from `ExternalDependencyReference.Kind`.
- Always emitted as an array, possibly empty.

### Deliberately not modelled

- **No "UI glue" axis.** Form-layout analysis is out of scope. UI handlers surface naturally as `eventHandler + writesState + readsExcelObjectModel` — accurate, just not labelled "UI."
- **No cyclomatic complexity.** Needs a deeper parser. Defer; revisit if hints feel too coarse against real workbooks.

## Coupling block — computation rules

Single pass over v1's `CallEdge[]`, scoped whole-workbook (always; ignores `moduleName`). All counts dedupe edges by `(fromModule, fromProc, toModule, toProc)` — repeated calls in a loop count once. Unresolved edges (dynamic dispatch, late binding) are excluded so they don't inflate `Ce`.

Definitions per call edge `(fromModule.fromProc → toModule.toProc)`:
- *Internal edge:* `fromModule == toModule`.
- *External edge:* `fromModule != toModule`.

`moduleCoupling[]` — one entry per module that appears in the workbook (read from `Modules[]`, not from edges, so isolated modules still get a zero-row):
- `ca` (afferent) — count of distinct external edges where `toModule == M`.
- `ce` (efferent) — count of distinct external edges where `fromModule == M`.
- `instability` — `ce / (ca + ce)` as a `double`. When `ca + ce == 0`, emit `0.0` (a module with no inter-module traffic is maximally stable).
- `internalEdges` — distinct internal edges within M.

`couplingPairs[]` — `Dictionary<(fromModule, toModule), int>` over external edges, projected to `[{from, to, edgeCount}]`. Sorted descending by `edgeCount`, then ascending by `from`, then by `to` for stable output. Pairs with `edgeCount == 0` are omitted.

Direction matters: `(A → B, 7)` and `(B → A, 2)` are separate entries — agents need that asymmetry to reason about extraction order ("A depends on B more than B depends on A → extract B first").

A second pass over the edge list yields no benefit. `ca/ce/internalEdges/couplingPairs` are all locally aggregable from the edge stream; `instability` is a finalization step over the per-module counters (O(modules), not O(edges)). The shapes that *would* need a second pass — Louvain modularity, normalised edge weights — are not in scope.

### Output size posture (Air.xlsm — 107 modules)

- `moduleCoupling`: 107 entries × ~50 bytes ≈ 5KB.
- `couplingPairs`: bounded by non-zero off-diagonal entries; ~25KB worst case.
- Total coupling block stays well below v2's `graph_too_large` ceiling.

## `targetParadigm` overlay — emission rules

When the caller passes `targetParadigm`, every emitted procedure gets a populated `csharpSuggestion`. Mapping is rule-based on the axes; first matching row in the paradigm's table wins.

### Common naming (paradigm-independent)

- `suggestedClassName` — module name PascalCased with VBA prefixes stripped (`mod` / `cls` / `frm`). Examples: `modOrders` → `Orders`, `clsCustomer` → `Customer`, `Module1` → `Module1` (passes through if no prefix).
- `suggestedMethodName` — procedure name PascalCased: `processOrder` → `ProcessOrder`, `do_thing` → `DoThing`.
- `isPublic` — mirrors VBA `IsPublic`.

### `classLibrary`

| Axes match | `targetType` | `lifetime` | `blockers` |
|---|---|---|---|
| `purity == pure` AND `shape == leaf` | `staticMethod` | `static` | — |
| `purity ∈ {pure, readsState}` AND `dependencies` empty | `staticMethod` | `static` | — |
| `purity == sideEffectful` AND `dependencies` includes `database`/`network` | `instanceMethod` | `scoped` | `requires_external_dependency_injection` |
| `purity == writesState` AND `dependencies` includes `excelObjectModel` only | `instanceMethod` | `scoped` | `depends_on_excel_object_model` |
| `trigger == eventHandler` | `requiresManualReview` | `null` | `event_handler_no_pure_classlib_target` |

### `workerService`

| Axes match | `targetType` | `lifetime` | `blockers` |
|---|---|---|---|
| `trigger == macroEntryPoint` AND `purity ∈ {writesState, sideEffectful}` | `backgroundService` | `singleton` | — |
| `trigger == eventHandler` AND name matches `Workbook_Open` / `Auto_Open` / `OnTime` patterns | `backgroundService` | `singleton` | — |
| any other procedure | `instanceMethod` | `scoped` | (collaborator method) |
| `dependencies` includes `excelObjectModel` (any row above) | (above) | (above) | append `depends_on_excel_object_model` |

### `webApi`

| Axes match | `targetType` | `lifetime` | `blockers` |
|---|---|---|---|
| `trigger == macroEntryPoint` AND `isPublic == true` | `apiAction` | `scoped` | — |
| any other procedure | `instanceMethod` | `scoped` | (helper for an action) |
| `dependencies` includes `excelObjectModel` (any row above) | (above) | (above) | append `depends_on_excel_object_model` |

### `console`

| Axes match | `targetType` | `lifetime` | `blockers` |
|---|---|---|---|
| `trigger == macroEntryPoint` | `consoleEntryPoint` | `null` | — |
| any other procedure | `staticMethod` if `purity ∈ {pure, readsState}`, else `instanceMethod` | `static` / `scoped` | — |

### Blocker codes — closed set, stable identifiers

- `depends_on_excel_object_model` — needs Interop or removal of Excel coupling.
- `requires_external_dependency_injection` — DB/network dep needs a registered service.
- `event_handler_no_pure_classlib_target` — Excel event semantics don't translate to a class library.
- `mutates_global_state` — touches module-scope variables; refactor before extraction.
- `unresolvable_call` — calls procedures not in the workbook (`Application.Run`, late binding); agent will need to follow up.

Multiple blockers can apply; emitted as an array.

### Rationale field under the overlay

When `targetParadigm` is set, the rationale gets one extra concrete sentence appended explaining the paradigm choice, e.g. "Suggested as `staticMethod` on `Orders` because pure + leaf + no Excel dependencies." Concrete enough to act on without re-reading the axes.

## Errors

**One new code:**
- `unsupported_paradigm` — `targetParadigm` value is not in `{classLibrary, workerService, webApi, console}`. Message lists supported values. Errors rather than silently falling through to `requiresManualReview` for everything.

**Inherited from v1:**
- `file_not_found`, `invalid_path` — from `PathGuard`.
- `parse_error` — from analyzer (corrupted VBA project).
- `module_not_found` — when `moduleName` doesn't resolve; reuses v1's existing helper, including the "available modules: …" suffix.

No `procedure_not_found` (v3 doesn't take a procedure parameter); no `graph_too_large` (v3's payload is bounded by procedure count, not edge count).

## Performance

| Step | Air.xlsm budget | Notes |
|---|---|---|
| v1 analyzer call | ~115ms | already measured |
| Axes pass — per procedure | ~5ms | 200 procs × constant work each |
| Coupling pass — per edge | ~3ms | 938 edges × counter increments |
| Overlay pass — per hinted procedure | ~5ms | table lookup on axes |
| **Total target** | **< 200ms** | with `targetParadigm` set; without, < 150ms |

Measured via `Stopwatch` and reported in `summary.wallTimeMs`. Mirrors v1's pattern.

## Testing

Pattern matches v1 (`VbaSourceAnalyzerTests`) and v2 (`VbaCallgraphFilterTests`).

### Unit tests — `tests/mcpOffice.Tests/Excel/Vba/VbaConversionHintBuilderTests.cs`

1. **Axis rules** — one fact per row above (~12 tests). Synthetic `Procedure` / `CallEdge` / `ObjectModelReference` records, no real `.xlsm` needed. Asserts the `axes` object exactly.
2. **Coupling computation** — separate `CouplingComputerTests.cs`. Hand-crafted edge lists; assert `ca/ce/instability/internalEdges` per module, then `couplingPairs` ordering and content. Includes the "all zero → instability == 0.0" edge case.
3. **Paradigm overlay** — one fact per row in the matrices (~15 tests). Same synthetic-record approach, asserts `csharpSuggestion` fields including `blockers[]`.
4. **Filter semantics** — `moduleName` filters `procedureHints[]` (asserted) but does *not* filter `moduleCoupling` / `couplingPairs` (also asserted).
5. **Naming convention** — `mod` / `cls` / `frm` prefix stripping for `suggestedClassName`.

### Real-world benchmark — gated, mirrors `AirSampleAnalysisTests`

`tests/mcpOffice.Tests/Excel/Vba/AirSampleConversionHintsTests.cs` against `C:\Projects\mcpOffice-samples\Air.xlsm`. Skips when absent.

Asserts:
- Total procedures equals analyzer's count (200).
- Every procedure has a hint.
- Coupling block has 107 modules.
- Wall time < 200ms.
- One smoke assertion per paradigm: `targetParadigm: "classLibrary"` produces at least one `staticMethod` and at least one `requiresManualReview`. Confirms the matrix actually fires under realistic data.

### Synthetic-fixture test — unconditional

`tests/mcpOffice.Tests/Excel/Vba/SyntheticConversionHintsTests.cs`. Reuses `tests/fixtures/synthetic-vba.xlsm`. End-to-end through `ExcelWorkbookService.SuggestVbaConversion`.

### Integration test — `tests/mcpOffice.Tests.Integration/`

One happy-path round-trip via stdio: synthetic fixture, call `excel_suggest_vba_conversion`, assert response shape and one rule-fired field. Don't re-test the matrix through stdio — unit tests cover that.

### Tool-surface test

`ToolSurfaceTests.cs` updated to include `excel_suggest_vba_conversion` (count 25 → 26).

### Error tests

- `unsupported_paradigm` triggered with `"blazor"` — assert `[unsupported_paradigm]` prefix and the supported-values list in the message.
- `module_not_found` triggered with a bogus `moduleName` — same assertions as v1.

## Architecture fit

```
ExcelWorkbookService.SuggestVbaConversion(path, moduleName?, targetParadigm?)
        |
        +--> VbaSourceAnalyzer.Analyze(path, moduleName?)        // existing v1 path
        |        returns VbaProjectAnalysis
        |
        +--> VbaConversionHintBuilder.Build(analysis, targetParadigm?)
                |
                +-- AxisClassifier            -- per-procedure axes
                +-- CouplingComputer          -- moduleCoupling + couplingPairs
                +-- ParadigmOverlayApplier    -- csharpSuggestion + rationale appendage
                |
                returns ConversionHints
```

New files under `Services/Excel/Vba/`:
- `VbaConversionHintBuilder.cs` — orchestrator.
- `AxisClassifier.cs` — pure function over a procedure + analysis.
- `CouplingComputer.cs` — pure function over the call graph.
- `ParadigmOverlayApplier.cs` — pure function over axes + paradigm.

New tool method on `Tools/ExcelTools.cs`: `ExcelSuggestVbaConversion` — one-line delegate to the service. New service interface method on `IExcelWorkbookService`.

New DTOs under `Models/`: `ConversionHints`, `ProcedureHint`, `ProcedureAxes`, `CSharpSuggestion`, `ModuleCoupling`, `CouplingPair`. All records, file-scoped namespace, nullable enabled.

## What this design deliberately does not do

- **No cluster detection** (Louvain or otherwise). Pairwise edge weights are the substrate; clustering is a follow-up that can layer on top.
- **No pagination on `procedureHints[]`.** Filed under the same TODO as the analyzer's heavy-array pagination — solved once for both.
- **No LLM-assisted hint refinement.** v3 is deterministic and cheap; the LLM consumer can layer reasoning on top.
- **No `blazor` / `winforms` / `wpf` paradigms.** Need form-layout analysis; defer.
- **No cyclomatic complexity per procedure.** Needs a deeper parser.
- **No "UI glue" axis.** Same reason — would require form analysis.
- **No conversion-confidence score.** Every hint is the rule-fire output; agents and humans can apply their own confidence based on the blockers array. A numeric score risks being treated as ground truth.

## Open questions deferred to implementation

- **Module-scope-write detection regex.** The `purity` axis distinguishes `readsState` vs `writesState` partly via a regex over the procedure's source range looking for `=` to module-scope identifiers. Need to confirm the regex is accurate enough on Air.xlsm without false positives (e.g. local variables shadowing module-level names). Spike during the axis-classifier task; fall back to `readsState` conservatively if the heuristic is noisy.
- **Macro-entry-point detection.** Procedure with public visibility + zero callers + non-`documentModule` kind. Need to verify on Air.xlsm that this matches what a human would call "entry point" — risk is over-reporting orphans.
- **`ListSheets` / `Worksheets` indexer quirk.** Already worked around in v1 via `MaterializeWorksheets`. Confirm v3 doesn't trip a similar issue when reading module names; if it does, lift the same workaround.

## Where to look for what

| You want to know... | Look at |
|---|---|
| Why v3 looks the way it does | this file |
| How v3 was built (TDD task list) | `docs/plans/2026-05-07-mcpoffice-excel-analyze-vba-v3-plan.md` (next) |
| v1 design (analyzer foundation) | `docs/plans/2026-05-03-mcpoffice-excel-analyze-vba-design.md` |
| v2 design (renderer — same architectural pattern) | `docs/plans/2026-05-03-mcpoffice-excel-render-vba-callgraph-design.md` |
| Current branch / next step | `SESSION_HANDOFF.md` |
| Project conventions | `CLAUDE.md` |
| Codebase shape | `ARCHITECTURE.md` |
