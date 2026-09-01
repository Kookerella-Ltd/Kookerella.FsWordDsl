# DSL ↔ OOXML mapping

Kookerella.FsWordDsl's DSL (`src/Kookerella.FsWordDsl/{Styles,NamedStyles,Numbering,Hyperlinks,
Protection,PageSetup,Tables,Images,Model}.fs`) aims to map 1:1 onto WordprocessingML wherever
the DSL models a feature at all. This document lists every place where that mapping is
inexact, lossy, or simply not implemented yet, so you know what to expect from a round trip
(`Document.save` then `Document.load`) and what would need to be added to close the gap.

## Modeled faithfully (1:1 or as close as makes sense)

- **Paragraphs and runs**: a `Paragraph`'s `Inlines` hold several independently-styled
  `Run`s - unlike Excel's own `CellValue` (always one uniformly-styled `Text` cell), rich
  text (mixed formatting within one paragraph) is first-class here, not a documented gap.
  Run formatting: font family, size, bold, italic, underline (six named styles plus an
  `OtherUnderline` escape hatch), strikethrough, color, highlight (Word's own fixed
  16-color palette, not arbitrary RGB), superscript/subscript.
- **Paragraph formatting**: alignment, spacing before/after, line spacing (single/1.5/
  double/at-least/exactly/multiple), indentation (left/right/first-line/hanging),
  keep-with-next, page-break-before.
- **Named styles** (`styles.xml`): paragraph and character styles, with `BasedOn`
  inheritance (the chain itself isn't resolved by this DSL - see the gap below) and a
  small built-in catalog (`BuiltInStyles.normal`/`heading1`/`2`/`3`/`title`/
  `listParagraph`/`hyperlinkCharStyle`) with explicit formatting rather than relying on
  Word's own built-in template defaults.
- **Numbered/bulleted lists**: single-level list definitions (bullet glyph + font, decimal,
  lower/upper letter, lower/upper roman, plus `OtherFormat` for any other raw OOXML
  `numFmt`), indentation, start-at value. This DSL collapses WordprocessingML's own
  abstract-numbering/numbering-instance indirection into one `NumberingDefinition` per
  distinct list - see the gap below on multi-level lists.
- **Tables**: rows and cells with horizontal merge (`GridSpan`) and vertical merge
  (`Restart`/`Continue`), per-cell shading and border overrides, column widths, a table
  style reference (name + banding flags - not a style *definition*, see the gap below),
  table-level borders (outer sides plus inside horizontal/vertical). A cell's content is a
  `Block list`, so a cell containing a nested table is exactly how Word itself represents
  one.
- **Images**: PNG/JPEG/GIF/BMP raster images anchored inline within a run (the natural
  placement for Word, unlike Excel's cell-range anchor) - `Data` is the image file's own
  raw bytes, embedded and read back byte-for-byte with no decoding/re-encoding.
- **Hyperlinks**: external (any URL, including `mailto:`) and internal (same-document
  bookmark) targets, wrapping one or more runs, with an optional tooltip.
- **Bookmarks**: named ranges within the document, wrapping inline content - scoped to
  within a single paragraph (see the gap below).
- **Comments**: modern Word comments (author, initials, date, text) anchored to a range of
  inline content - scoped to within a single paragraph (see the gap below). Not the legacy
  "cell comment" concept Excel models; this is what current Word's UI calls a comment.
- **Simple fields**: raw field instruction text (e.g. `"PAGE"`) plus a cached display value
  - this DSL never evaluates a field itself, the same "cachedValue is the only number that
    will ever exist until something else computes one" posture Excel's `CellValue.Formula`
    documents; real Word recalculates on open and overwrites it.
- **Sections and page setup**: a document is a sequence of `Section`s, each with its own
  page size (named set plus custom/other escape hatches), orientation, margins, column
  count, starting page number, and break type (`SectionBreakType`: next-page - the default,
  not written - continuous, even-page, odd-page).
- **Footnotes and endnotes**: `Inline.Footnote`/`Endnote` mark a point in a paragraph's own
  `Inlines`; `content` is the note's own body (`Block list` - ordinary paragraphs, or a
  table) written to `word/footnotes.xml`/`endnotes.xml` and given a document-scoped
  sequential id automatically. `Writer` prepends the note-reference-mark run itself (`w:
  footnoteRef`/`w:endnoteRef`) to the body's first paragraph, and both parts always carry
  the two boilerplate separator/continuation-separator entries a real Word-authored file
  has - a caller never has to think about either. See the gap below on custom numbering.
- **Headers and footers**: `Default`/`First`/`Even` variants per section, each arbitrary
  `Block` content - the `titlePg`/`evenAndOddHeaders` flags real Word needs to actually
  honor `First`/`Even` are set automatically whenever they're used.
- **Document protection**: a single edit restriction (read-only/comments-only/
  tracked-changes-only/forms-only) plus an optional password, using the modern salted,
  iterated SHA-512 hash scheme (see the gap below on verification).
- **Macros (VBA)**: `Document.VbaProject` embeds a real `word/vbaProject.bin` and switches
  the saved file's package content type to Word's macro-enabled kind so Word actually
  trusts and runs it - like Excel's own `Workbook.VbaProject`, this DSL does no
  encoding/decoding of VBA source itself, only embeds and hands back exactly the bytes
  it's given.
- **Code generation**: `Document.generateScript` renders any `Document` value (including
  one produced by `Document.load`) back out as an `.fsx` script that rebuilds a
  structurally equivalent file when run via `dotnet fsi` - verified for every scenario
  under `tests/Kookerella.FsWordDsl.Tests/Examples/` by the `Category=Slow` test group,
  which actually executes each one, not just generates it. Unlike Excel's own `CodeGen`,
  this always renders every field explicitly rather than diffing against each type's
  `.Default` - simpler, more verbose output; a good target for a future pass.
- **XML and JSON surfaces**: `Xml.toDocument`/`Xml.ofDocument` (against the embedded
  `Xml.xsd`) and `Json.toDocument`/`Json.ofDocument` (against the test-suite-only
  `Json.schema.json`) cover the same feature set as the rest of this library, for a caller
  who'd rather generate or consume data than write F#/C# at all.

## Known gaps (documented, not silently "supported")

- **Track changes.** Insertions, deletions, and other revision marks aren't modeled at
  all - a document is always written (and read) as if "accept all changes" had already
  been applied.
- **Content controls (structured document tags).** Not modeled.
- **Footnote/endnote numbering customization.** Always continuous decimal numbering
  starting at 1, document-wide (Word's own default) - a custom number format, or
  restarting the count per section/page, isn't modeled (`w:footnotePr`/`w:endnotePr` at the
  section level aren't written or read).
- **Real field computation.** Only raw instruction text + a cached display value round-trip
  (see "Modeled faithfully" above) - a table of contents, cross-reference, or any other
  field that depends on document layout is never actually computed by this DSL, unlike
  Excel's pivot tables (which *do* perform real aggregation at write time).
- **Comments/bookmarks spanning more than one paragraph.** Both are scoped to wrapping
  inline content within a single paragraph in this DSL - a foreign file with either
  spanning multiple paragraphs degrades on read (the range is not reconstructed across the
  paragraph boundary).
- **Table style *definitions*.** `TableStyleRef.Name` is a reference to a style by name (a
  built-in like `"TableGrid"`, or a custom one defined elsewhere in the document) - this
  DSL doesn't model custom table style *definitions* themselves, only the reference, same
  documented gap Excel's own `TableStyle.Name` has.
- **Multi-level numbering.** `NumberingDefinition.Levels` supports several levels, but this
  DSL doesn't validate that a paragraph's `(numId, level)` reference and a level's own
  `Text` placeholder pattern (e.g. `"%1.%2"`) actually agree - malformed combinations write
  and read back verbatim rather than being caught.
- **Theme colors.** `Styles.Color` only models `Rgb` and `Auto` - a run or highlight using
  a theme color reference isn't modeled, same documented gap Excel's own unresolved
  `Theme`/`Indexed` colors have.
- **Text boxes, SmartArt, embedded charts/OLE objects.** Not modeled - a Word document's
  drawing canvas is used here only for the one inline-image case (see "Modeled faithfully"
  above).
- **Digital signatures.** Not modeled.
- **Password verification against real Word.** The modern salted-iterated-SHA512 scheme is
  implemented per the published algorithm, but unverified against real Word (no Word
  available in this environment to confirm acceptance) - treat this with the same "verify
  separately before relying on it" caution Excel gives its own Sparklines feature.
- **VBA project authoring, and non-default sheet/document codenames.** Same gap Excel's own
  `Workbook.VbaProject` documents: this DSL embeds and reads back a `vbaProject.bin`
  byte-for-byte but never inspects, decompiles, or generates its contents, and doesn't
  model a caller-chosen internal codename.
- **No real-world macro test asset.** Unlike Excel's `VbaMacro` scenario (which round-trips
  against a `vbaProject.bin` actually extracted from a file Excel itself saved), this
  repo's `Macro` scenario uses synthetic bytes - no real Word-produced macro project was
  available in this environment to substitute. The round-trip mechanics are identical
  either way (opaque byte passthrough), but this specific scenario hasn't been verified
  against a real compiled VBA project the way Excel's has.

## Out of scope for this pass

Nothing is currently in this bucket in the sense Excel's own `MAPPING.md` uses it (every
SpreadsheetML feature Excel originally scoped out has since been implemented) - this repo
is a first pass (F# core only; no C# wrapper or MCP server yet, see `CLAUDE.md`), so
several of the "Known gaps" above (track changes, content controls, text boxes) are
realistic candidates for a genuine future extension rather than permanently excluded.

## A note on formatting

Unlike Excel (where `CellStyle` needs interning into a shared, indexed stylesheet - two
cells with structurally equal styles get deduplicated into one `cellXfs` entry),
WordprocessingML writes direct run/paragraph formatting inline on every element's own
`w:rPr`/`w:pPr` - there is no equivalent index, so this DSL doesn't intern or deduplicate
direct formatting the way Excel does. Named styles (`styles.xml`) are the one place this
DSL's model and Excel's converge: neither DSL resolves a style's inheritance chain itself,
both just ensure every referenced style id gets a real definition written.
