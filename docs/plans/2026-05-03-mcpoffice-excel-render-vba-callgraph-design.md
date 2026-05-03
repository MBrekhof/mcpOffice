# mcpOffice — `excel_render_vba_callgraph` Design

**Date:** 2026-05-03
**Status:** Approved (brainstorming phase)
**Scope:** A new MCP tool that emits the VBA call graph as Mermaid or DOT for visual inspection. Layered on top of the existing `excel_analyze_vba` analyzer. Conversion hints and coupling-score features are out of scope here — they remain on the v3 / v4 roadmap.

## Purpose

Help an agent answer the question *"how is this workbook organised — what calls what?"* with a picture rather than 938 JSON edges.

The `excel_analyze_vba` JSON output already carries the call graph as data. That's good enough for programmatic consumption, but it's the wrong shape when:

- The agent wants to render the graph inline (Mermaid in Claude's chat UI).
- A human needs to skim a workbook's structure and decide what to dig into.
- The visual cluster-by-module view tells the migration story faster than reading edges as text.

A separate render tool keeps that concern out of the analyzer (which stays focused on emitting structural facts) while reusing every byte of the analyzer's parsing work.

## Architecture

`excel_render_vba_callgraph` reuses the analyzer pipeline, then layers a filter and a renderer:

```
excel_render_vba_callgraph (Tool)
  ↓
ExcelWorkbookService.RenderVbaCallgraph(path, options)
  ↓
VbaProjectReader.Read(path)            ← existing
  ↓ ExcelVbaProject { modules[] }
VbaSourceAnalyzer.Analyze(project, ...) ← existing (call graph + procedure index)
  ↓ ExcelVbaAnalysis (in-memory only — never serialised here)
VbaCallgraphFilter.Apply(analysis, options)  ← new
  ↓ FilteredGraph { nodes[], edges[] }
ICallgraphRenderer.Render(filtered)          ← new (interface)
   ├── MermaidCallgraphRenderer
   └── DotCallgraphRenderer
  ↓ string
```

Two new units, one new interface:

- **`VbaCallgraphFilter`** under `Services/Excel/Vba/`. Pure function: takes the analyzer's full graph plus filter options, returns a `FilteredGraph` with surviving nodes/edges and a node-cap verdict. No I/O, no rendering. Independently testable against in-memory analyses.
- **`ICallgraphRenderer`** plus two implementations under a new `Services/Excel/Vba/Rendering/` folder. Each impl takes a `FilteredGraph`, returns a string. Strict separation: the filter knows nothing about output syntax; the renderers know nothing about VBA.

The split exists because the filter logic — BFS, direction handling, cluster expansion — is non-trivial and identical across formats. Testing it once buys correctness for both renderers.

`excel_analyze_vba` is unchanged. This feature is purely additive.

## Tool surface

```
excel_render_vba_callgraph(
    path,                       // absolute path to .xlsm / .xlsb
    format = "mermaid",         // "mermaid" | "dot"
    moduleName = null,          // optional: scope to one module
    procedureName = null,       // optional: focal procedure (requires moduleName)
    depth = 2,                  // BFS hops from focal procedure
    direction = "both",         // "callees" | "callers" | "both"
    layout = "clustered",       // "clustered" | "flat"
    maxNodes = 300              // hard cap; throws graph_too_large past this
)
```

Returns a single string — the rendered graph. No JSON wrapper. The agent already knows from its own call which format it asked for, and the rendered text is self-describing.

### Filter modes

Three modes emerge from the argument combinations:

1. **No filters** → render the whole workbook clustered. Most workbooks fit under 300 nodes; large ones (e.g. Air.xlsm) trip `graph_too_large` and the agent has to narrow.
2. **`moduleName` only** → all procedures in that module plus their direct neighbours (one hop out, both directions). The "what does this module talk to?" view.
3. **`moduleName + procedureName`** → focal-procedure BFS using `depth` (hops out) and `direction` (callees / callers / both). The surgical view that makes a 938-edge workbook readable.

`procedureName` without `moduleName` is rejected with `invalid_render_option`. Bare procedure names aren't unique — every sheet with a `Worksheet_Change` would match — and silently picking one is worse than refusing.

### Layouts

- **`clustered`** (default) — wrap each module's procedures in a Mermaid `subgraph` / DOT `cluster`. Module boundaries become visible at a glance, which is almost always what the consumer wants.
- **`flat`** — every procedure as a top-level node, FQN as label. Useful when the cluster boxes get in the way (very small graphs, or when only one module survives the filter).

## Visual conventions

This is where "should provide insight" cashes out. Three node classes, two edge classes, rendered identically across Mermaid and DOT.

### Node classes

- **Event handler** (`isEventHandler = true`) — rounded rectangle, light-blue fill. These are the workbook's entry points; the eye should find them first.
- **Orphan** — no callers, no callees, *and* not an event handler. Dashed border. Often dead code or ribbon-button targets that VBA can't see; worth flagging.
- **External** — placeholder node for an unresolved callee (e.g. `Application.Run "Foo"` where `Foo` isn't in the workbook, or a builtin like `MsgBox`). Grey fill, dashed border.
- Everything else: default rectangle.

### Edge classes

- **Resolved call** — solid arrow.
- **Unresolved call** — dashed arrow into an `<external>` node.

No edge labels. Line numbers add visual noise; the analyzer JSON already carries them when needed.

### Node identity & label

- Node ID is the FQN (`Module.Proc`) — stable, unique, valid in both formats after escaping.
- Label inside the node:
  - `clustered` mode: bare procedure name (the cluster header carries the module).
  - `flat` mode: `Module.Proc`.
- `<external>` node label is the verbatim unresolved callee name.

### Mermaid example (sketch, ScreeningDB-V2 partial)

```
flowchart TD
  subgraph mdlScreeningDB
    ReadExports[ReadExports]
    ReadExportsLC([ReadExportsLC]):::handler
    SaveDB[SaveDB]
    Paste2Cell[Paste2Cell]
  end
  subgraph Blad17
    Blad17_Worksheet_BeforeDoubleClick([Worksheet_BeforeDoubleClick]):::handler
  end
  ReadExports --> ReadExportsLC
  ReadExports --> SaveDB
  Blad17_Worksheet_BeforeDoubleClick --> Paste2Cell
  classDef handler fill:#e1f5ff,stroke:#0277bd
  classDef orphan stroke-dasharray:5 5
  classDef external fill:#f5f5f5,stroke-dasharray:3 3
```

DOT mirrors the same conventions via `subgraph cluster_*`, node `shape` / `fillcolor`, edge `style="dashed"`.

## Filter semantics — details

### Direct-neighbour expansion (`moduleName` only)

Given module `X`:

1. Seed nodes = every procedure in module `X`.
2. For each seed, add every direct caller and every direct callee (one hop, both directions).
3. Surviving edges = edges whose endpoints are both in the surviving node set.

External (unresolved) callees become `<external>` nodes only when their resolved-side endpoint survives.

### External-node identity

One `<external>` node per distinct unresolved callee name. `MsgBox`, `Application.Run`, `CreateObject` each get their own node, deduplicated across the whole filtered graph. Edges from multiple resolved procedures into the same external converge on the same node — this is what makes "everything calls `MsgBox`" visually obvious rather than scattered.

### Focal-procedure BFS (`moduleName + procedureName`)

Given focal procedure `X.Y`, `depth = N`, `direction = D`:

1. Seed = `{X.Y}`.
2. Expand BFS up to `N` hops. At each hop:
   - `direction = "callees"` — follow outgoing edges only.
   - `direction = "callers"` — follow incoming edges only.
   - `direction = "both"` — follow both.
3. Surviving edges = same rule as above.

`depth = 0` is legal — renders just the focal node and its module's cluster header.

### Orphan classification

A node is an orphan when **after filtering** it has zero in-edges, zero out-edges, and `isEventHandler = false`. The classification is per-render, not per-workbook — a procedure could be visually an orphan in one filtered view and central in another.

### `maxNodes` cap

Counted after filtering, before rendering. If exceeded, throw `graph_too_large` immediately — no truncation. The error message includes the actual node count and suggests the next narrowing step (`add moduleName`, `add procedureName`, `reduce depth`).

The default of `300` is a Mermaid-readability budget. Empirical: graphs past ~300 nodes / ~500 edges produce diagrams that a human can't read inline and that Claude's UI starts to render slowly. The agent can pass a higher cap if it's piping into DOT for offline rendering.

## Error model

Same fast-failure tier as `excel_analyze_vba` (file / path / lock / parse), plus three render-specific codes:

| Code | When |
|---|---|
| `file_not_found`, `invalid_path`, `vba_project_locked`, `vba_parse_error` | unchanged — same set as `excel_extract_vba` and `excel_analyze_vba` |
| `module_not_found` | reused — `moduleName` doesn't match any module. Message lists available module names |
| `procedure_not_found` | new — `procedureName` doesn't match any procedure in `moduleName`. Message lists candidates |
| `graph_too_large` | new — node count exceeds `maxNodes`. Message includes actual count and a narrowing suggestion |
| `invalid_render_option` | new — unknown `format` / `layout` / `direction` value, or `procedureName` passed without `moduleName`. Lists valid options |

`hasVbaProject = false` is **not** an error. Returns an empty rendered graph (`flowchart TD\n` for Mermaid, `digraph G { }\n` for DOT) so the agent sees the no-macros state explicitly.

## Testing

Three layers, mirroring v1.

### Unit — `VbaCallgraphFilter` against synthetic analyses

Build in-memory `ExcelVbaAnalysis` objects (no `.xlsm` involved) and assert filter behaviour:

- No filters → all nodes/edges survive.
- `moduleName = X` → exactly the X-module procedures plus direct neighbours; non-neighbours dropped.
- `moduleName = X, procedureName = Y, depth = 2, direction = "callees"` → BFS frontier matches by hand.
- `direction = "callers"` — mirror of the above.
- `direction = "both"` — union.
- Orphan classification — procedure with zero in/out edges flagged orphan; same procedure ceases to be orphan when `isEventHandler = true`.
- `maxNodes` — returns "too large" verdict; never truncates.
- `module_not_found` / `procedure_not_found` — filter throws with the expected code.

### Unit — renderers against synthetic `FilteredGraph` objects

Drive each renderer with hand-built filtered graphs:

- One node, no edges — emits valid empty-ish graph.
- One cluster with two procedures and one edge — subgraph syntax verified.
- Resolved + unresolved edges — solid vs. dashed.
- Event-handler / orphan / external classes applied to the right nodes.
- Reserved characters in procedure names escaped correctly per format. VBA allows `[Bracketed Names]`; both renderers must produce parseable output. Test against a procedure whose name contains characters that are syntax in the target format.
- Mermaid structural assertions: every `subgraph` has a matching `end`; every node ID is referenced before use.
- DOT structural assertions: balanced braces; no unquoted IDs that contain reserved characters.

We do *not* shell out to `dot` or to a Mermaid renderer in tests — the structural assertions are sufficient and keep tests fast and dependency-free.

### Real-world benchmark

A `[Fact]` against `C:\Projects\mcpOffice-samples\Air.xlsm`, gated on file existence (same skip-when-absent pattern as the existing analyzer benchmark).

- Whole-workbook render (no filters) → `graph_too_large`. Confirms the cap is the safety net it's meant to be.
- `moduleName = <one of the 107 modules>` → renders successfully, well under `maxNodes`.
- `moduleName + procedureName + depth = 1` → small focal graph; line-count snapshot in a sensible band.
- Wall-time assertion — full pipeline (parse + analyze + filter + render) completes well under 500 ms. The analyzer alone is ~115 ms, so this is loose.

### Stdio integration

One test in `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs` — call `excel_render_vba_callgraph` against `tests/fixtures/sample-with-macros.xlsm`, assert the response is a non-empty string starting with `flowchart` (Mermaid default). Proves the protocol layer doesn't drop the string.

### Tool surface

Update `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs` to include `excel_render_vba_callgraph` (24 → 25 tools).

## Out of scope for v2

- **Conversion hints** — opinionated suggestions like "this VBA Sub maps to a C# class with method X". Stays on the v3 roadmap; will need real-world evidence first.
- **Cross-module coupling score** — quantitative refactoring guidance. v4 candidate.
- **Excalidraw output** — covered for free by piping the Mermaid output into the existing Excalidraw MCP server's `create_from_mermaid`. No first-class output format here.
- **DevExpress `DiagramControl` rendering** — UI control, would force a Windows message loop into a stdio console app and break the inline-Mermaid use case. If interactive editing becomes a real need, build it as a separate downstream consumer, not inside this server.
- **Object-model overlay** — rendering "this procedure touches Worksheets("X") and Range("Y")" as side annotations. The analyzer already surfaces this in its references array; visualising it is a different feature.
- **Time-budget cap.** No timeout. The analyzer runs in ~115 ms on Air.xlsm; rendering is text concatenation. If real-world workbooks blow that, add a budget later.
- **Caching across calls.** Each tool call re-parses. The agent's natural workflow is "narrow, narrow again" — caching the analyzer output across calls is a nice optimisation but optional. Punt to a follow-up.
- **`includeExternal` toggle.** External nodes are included by design (they're truthful — every `MsgBox` / `IsNumeric` / runtime-resolved call shows up). On VBA-builtin-heavy workbooks this can produce visual clutter. If it becomes a real annoyance, add a boolean toggle to suppress them; deferred until evidence of a need.
