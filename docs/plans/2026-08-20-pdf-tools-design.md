# PDF tools — design

Date: 2026-08-20. Branch `feat/pdf-tools`. Status: shipped.

## Why

mcpOffice could *write* PDFs (`word_convert` → `ExportToPdf`) but not read one. An agent handed a
PDF had nothing to call. The trigger was a batch of monthly lab reports ("Overzichtsrapport",
produced by the old ILIS LIMS) that needed reading: `Author = "WLN - ILIS"`,
`Creator = a2ps version 4.14`, `Producer = GPL Ghostscript 9.55.0` — i.e. a monospaced ASCII
report run through a2ps into PostScript and distilled to PDF. Columns carry the meaning, so
plain text extraction is not enough.

No new package was needed: `DevExpress.Document.Processor` already brings `DevExpress.Pdf.Core`
and `DevExpress.Pdf.Drawing` along.

## Tool surface

Seven tools, `pdf_` prefix, all read-only (writing PDFs stays `word_convert`'s job):

| Tool | Purpose |
|------|---------|
| `pdf_get_metadata` | properties, version, permissions, bookmark count, per-page geometry |
| `pdf_read_text` | text per page; `preserveLayout` rebuilds the fixed-width grid |
| `pdf_read_layout` | positioned words or visual lines with boxes |
| `pdf_find_text` | matches with page + bounding box |
| `pdf_render_page` | page → png/jpg/bmp/gif/tiff |
| `pdf_extract_images` | embedded rasters → PNG, with page placement |
| `pdf_get_outline` | bookmark tree |

## Decisions

**Coordinates are flipped to a top-left origin.** DevExpress reports `PdfOrientedRectangle.Top`
in PDF user space, where Y grows upward from the bottom of the page. Every consumer of these
tools wants reading order, and "sort ascending" being *backwards* is the kind of thing that
produces a silently reversed document. The flip happens once, in the service; the DTOs document
it.

**Lines are inferred, not read.** PDF stores no lines — only glyph placements. `LineGrouper`
clusters words whose vertical centres fall within `0.6 × word height`. Tolerance is derived from
height rather than being a constant, so a 6pt table and a 24pt heading in the same document both
group correctly. The running centre is re-averaged as words join, so a line that drifts slightly
across the page does not split in the middle.

**`preserveLayout` is opt-in, not the default.** It needs the word cursor (a full document walk)
where `GetPageText` does not, so prose callers should not pay for it. For column reports it is
the difference between usable and not.

**`LayoutTextRenderer` estimates one character width per page** as the *median* of
`wordWidth / textLength`, not the mean — a handful of wide-glyph or one-character words would
otherwise skew the grid. Words are then placed at `round((x - originX) / charWidth)`. `originX`
is the leftmost word on the page, so a uniform left margin does not become leading whitespace.
When two words claim the same column the later one is pushed one space right: overlapping output
is acceptable, losing a value is not.

**The word cursor is document-wide.** `PdfDocumentProcessor.NextWord()` is forward-only over the
whole document with no per-page overload, so a `pageRange` filters *after* the walk rather than
skipping ahead. Verified against a real report: 875 words walked, terminates on null. It is one
pass either way; the alternative is no positioned text at all.

**Fixtures are generated, not committed.** `TestPdfDocuments` builds a `RichEditDocumentServer`
document and calls `ExportToPdf`, reusing the already-tested Word→PDF path rather than
introducing a PDF-authoring API just for tests. This follows the repo rule against committed
binary fixtures.

## Verified

- 50 unit tests + 2 stdio integration tests. Full suite 351 unit + 17 integration, 2 skipped.
- Against the three real ILIS reports: `preserveLayout` reproduces the column layout exactly;
  `pdf_find_text` returns all four same-page occurrences of "Legionella" with distinct boxes;
  `pdf_render_page` at 150 dpi produces a clean, watermark-free PNG under the DevExpress trial
  license.

## Not built

- **Table extraction.** Turning lines into cells needs column-boundary inference that is
  guesswork on anything but a ruled table. `pdf_read_layout` gives the honest primitive; a caller
  who knows the report shape can slice it.
- **OCR.** Nothing here reads a scanned page. `pdf_render_page` is the escape hatch — render it
  and look at it.
- **PDF writing / editing / form filling.** DevExpress supports all of it; no caller has asked.
