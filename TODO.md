# TODO

Pending work for mcpOffice. Maintained by the `/handoff` skill.

Completed milestones (Word POC, Excel POC, analyzer v1/v2/v3, export_csv, Markdig converter) are summarized in git history and `SESSION_HANDOFF.md`. Only open follow-ups live here.

## excel_suggest_vba_conversion (v3) — deferred follow-ups

- [ ] Cluster detection (Louvain) on the module graph; layer on top of pairwise coupling.
- [ ] Pagination on `procedureHints[]` for very large workbooks (same TODO as analyzer's heavy arrays).
- [ ] `blazor` / `winforms` / `wpf` paradigms — need form-layout analysis the regex layer can't reliably do.
- [ ] Cyclomatic complexity per procedure — needs a deeper VBA parser.
- [ ] Module-scope-write detection regex — currently `purity` collapses to 3 values (`pure` / `readsState` / `sideEffectful`); `writesState` activates when `ExcelVbaObjectModelRef.Mode` lands.
- [ ] Dependencies-axis schema drift: design's closed set said `{excelObjectModel, filesystem, database, network, registry, shell}` but v1's `VbaReferenceCollector` emits `file` (not `filesystem`) — observed on RingOnderzoek.xlsm. v3 currently only renames `automation → shell` and passes everything else through. Either rename `file → filesystem` in v3's mapping or change v1 to emit the design's spelling. Probably the latter (keeps v1's emissions intelligible to other consumers).
- [ ] `ParadigmOverlayApplier.StripModulePrefix` only handles `mod` / `cls` / `frm`. Real-world Air.xlsm uses `mdl` (e.g. `mdlAIR`, `mdlBalans`) — currently passes through as `MdlAIR`. Either extend the prefix list (`mdl`, `bas`, `srv`, etc.) or make it configurable per workbook. Surfaced via 2026-05-07 live verification.

## excel_export_csv — deferred follow-ups

- [ ] `excel_export_ndjson` sibling — column-typed output for `pandas.read_json(lines=True)` consumers. Shares streaming infrastructure with `excel_export_csv`.
- [ ] `.csv.gz` compression — wrap the `FileStream` in `GZipStream` when `outputPath` ends in `.gz`. Trivial follow-up; deferred for v1 to keep the test matrix small.
- [ ] Optional `lineEnding` (`crlf` / `lf`) and `delimiter` (`,` / `\t` / `;`) parameters — only if a real consumer surfaces a need. CSV/TSV agents don't typically need this.
- [ ] Optional bulk-export mode that loads the workbook once and writes N CSVs — surfaced by ScreeningDB sweep (each `ExportCsv` reloads, ~28s per call on the 26 MB workbook → 21 sheets = 10 minutes). Lower priority because most agent flows export 1-2 sheets per call.

## Markdig converter — deferred follow-ups

- [ ] **Unify `WriteCellInline` and `WriteInline`.** `WriteCellInline` mirrors all of `WriteInline`'s case logic with cursor-based anchoring instead of `para.Range.End`. New inline types currently need adding in both places. Refactor via a small writer abstraction (e.g. an `IInlineSink` with `Insert(text)` returning a `DocumentRange`) so the case logic lives in one place.
- [ ] **Normal-style polish for `word_create_from_markdown`.** Setting Calibri 11pt / 1.15 line spacing / 8pt SpacingAfter as document defaults would tighten body text. Skipped during the 2026-05-13 heading style pass because `MarkdownToDocxConverter.Apply()` is also called by `word_append_markdown` (mustn't fight an existing doc's Normal style). Either gate via a flag passed through the service, or split into `Apply` (append-safe) vs `ApplyToFreshDocument` (mutates defaults).

## Word tools — deferred follow-ups

- [ ] **`word_mail_merge` — add `overwrite` parameter.** (ID: 1367)
  It calls `PathGuard.RequireWritable(outputPath, overwrite: false)` with no override, so regenerating into the same path fails with `file_exists`; `word_create_blank` / `word_create_from_markdown` / `word_convert` already expose `overwrite=false`. Surfaced 2026-08-22 by the "menus as mail-merge renders" use case (CSV price list → `word_mail_merge` → `word_convert`, regenerated every run). Plan: `bool overwrite = false` through `IWordDocumentService.MailMerge`, the impl, and the tool parameter; one unit test merging twice to the same path. Tool name unchanged, so `ToolSurfaceTests` is untouched.

## PDF tools — deferred follow-ups

- [ ] **Table extraction (`pdf_extract_tables`).** Needs column-boundary inference on top of `LineGrouper` — clustering word x-positions across a page into column stops. Guesswork on unruled reports, which is why v1 stops at `pdf_read_layout`. Do it only when a caller has a report shape worth targeting.
- [ ] **OCR for scanned PDFs.** `pdf_read_text` returns empty for image-only pages; today's answer is `pdf_render_page` and look. Would need an external engine (Tesseract) — a real dependency decision, not a small addition.
- [ ] **`pdf_render_page` page ranges / contact sheet.** Currently one page per call. `CreateTiff(stream, pageNumbers, dpi)` already does multi-page in one shot if a caller wants it.
- [ ] **Per-page word cursor.** `NextWord()` is document-wide, so `pageRange` on `pdf_read_layout` filters after walking every page. Fine for reports, wasteful on a 500-page document. No DevExpress API for it — would need `GetText(PdfDocumentArea)` plus per-word search, which is likely slower, so measure before changing.
- [ ] **`pdf_extract_images` vector graphics.** Only raster images are extracted (`GetImagesInfo`). Charts drawn as vector paths are invisible to it; `pdf_render_page` is the fallback.
- [ ] **Text-extraction options.** `PdfTextExtractionOptions.ClipToCropBox` is not exposed. Only matters for documents with content outside the crop box.

## Side items

### Carried from Word POC
- [ ] Optional: baseline `.editorconfig` once enough files exist to enforce against.
- [ ] Add `[JsonDerivedType]` discriminators to the abstract `Block` record (and concrete `HeadingBlock`/`ParagraphBlock`) when tests start asserting on `word_read_structured`'s wire JSON.

### Carried from Excel POC
- [ ] PROJECTLCID-aware code page selection in `VbaProjectReader` (currently hardcoded to cp1252). MS-OVBA dir record `0x0002 PROJECTLCID` carries the project locale.
- [ ] `excel_get_structure`: optional pivot / chart / external-connection counts via Open XML walk (DevExpress doesn't expose them directly).
- [ ] `excel_list_formulas`: rough dependency-token extraction (deferred — formula text is enough for now).
- [ ] `VbaProcedureScanner` lacks tests for `ParamArray` parameter form and `Static Sub` procedure form. Both currently parse correctly per the regex; no behavior gap, just test coverage.
- [ ] **Pagination on `callGraph` and `references` arrays in `excel_analyze_vba`.** Even with a `moduleName` filter, the heaviest module on a large workbook can be too big. Add `offset` / `limit` (or cursor) to the heavy arrays so a caller can stream them. Lower priority than the module filter — the filter alone covers most real cases.
