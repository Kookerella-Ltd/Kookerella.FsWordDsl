# DSL ↔ OOXML mapping

Kookerella.FsWordDsl's DSL (`src/Kookerella.FsWordDsl/{Styles,NamedStyles,Numbering,Hyperlinks,
Protection,PageSetup,Tables,Images,Model}.fs`) aims to map 1:1 onto WordprocessingML wherever
the DSL models a feature at all. This document lists every place where that mapping is
inexact, lossy, or simply not implemented yet, so you know what to expect from a round trip
(`Document.save` then `Document.load`) and what would need to be added to close the gap.

## Reading a foreign file

Round-trip correctness is guaranteed and tested against this library's own `Writer` output
(the `verifyScenarioNamed` harness in `tests/Kookerella.FsWordDsl.Tests/Tests.fs`) - reading
an arbitrary real-world `.docx` (one Word itself produced, using features this DSL doesn't
model) is best-effort. Concretely, `Reader` never fails a whole document just because it
contains something unmodeled:
- Inside a paragraph, an unrecognized inline-level element (e.g. track-changes `w:ins`/
  `w:del` wrapping a run) is silently skipped - the surrounding recognized content still
  reads back fine, but that specific unmodeled content doesn't survive the read.
- At the document/cell body level, a content control (`w:sdt`) or `w:customXml` wrapper is
  unwrapped rather than skipped - the paragraphs/tables it wraps are read normally, only
  the wrapper itself (and whatever metadata it carried) is discarded. Anything else
  unrecognized there (`w:altChunk`, a bookmark/comment range marker placed directly in the
  body rather than nested in a paragraph, `w:permStart`/`permEnd`, ...) is dropped the same
  way the inline case is.
- `tests/Kookerella.FsWordDsl.Tests/Tests.fs`'s own `Reader tolerates unmodeled body-level
  content instead of throwing` test builds a file directly against the OOXML SDK
  (bypassing `Writer` entirely, since there's no DSL-level way to author these constructs)
  to exercise exactly this - it's the one test in the suite that isn't round-tripped
  through this DSL's own writer, by design.

## Modeled faithfully (1:1 or as close as makes sense)

- **Paragraphs and runs**: a `Paragraph`'s `Inlines` hold several independently-styled
  `Run`s - unlike Excel's own `CellValue` (always one uniformly-styled `Text` cell), rich
  text (mixed formatting within one paragraph) is first-class here, not a documented gap.
  Run formatting: font family, size, bold, italic, underline (six named styles plus an
  `OtherUnderline` escape hatch), strikethrough, color (`Rgb`, `Auto`, or a theme-relative
  `Theme` color - see below), highlight (Word's own fixed 16-color palette, not arbitrary
  RGB), superscript/subscript, small caps, all caps, hidden text.
- **Paragraph formatting**: alignment, spacing before/after, line spacing (single/1.5/
  double/at-least/exactly/multiple), indentation (left/right/first-line/hanging),
  keep-with-next, page-break-before, borders (`BorderStyle`, the same shape reused for
  table/cell borders - top/bottom/left/right only, no `between`/`bar`, same gap as table
  borders not modeling diagonals), shading (background fill color), custom tab stops
  (left/center/right/decimal/bar alignment, dot/hyphen/underscore/heavy/middle-dot
  leaders, plus an `OtherTabAlignment` escape hatch) - an empty `TabStops` list means "no
  custom tabs", not "clear Word's own inherited defaults" (this DSL doesn't author
  `w:val="clear"` entries).
- **Theme colors** (`Color.Theme`): Word's twelve standard theme slots (`ThemeColorKind`),
  with an explicit `Fallback` RGB and optional tint/shade - since this DSL has no theme
  part (`word/theme/theme1.xml`) to resolve a token against, real Word does the resolving,
  the same "always also write a computed value" convention Word itself follows so a
  themeless reader still sees something reasonable. Only modeled for run color and
  shading/fill backgrounds (`RunStyle.Color`, `ParagraphFormat.Shading`, `TableCellProps.
  Shading`, `TableStyleRegion.CellShading`) - see the gap below on border colors.
- **Named styles** (`styles.xml`): paragraph and character styles, with `BasedOn`
  inheritance (the chain itself isn't resolved by this DSL - see the gap below) and a
  small built-in catalog (`BuiltInStyles.normal`/`heading1`/`2`/`3`/`title`/
  `listParagraph`/`hyperlinkCharStyle`) with explicit formatting rather than relying on
  Word's own built-in template defaults.
- **Numbered/bulleted lists, including multi-level**: list definitions (bullet glyph +
  font, decimal, lower/upper letter, lower/upper roman, plus `OtherFormat` for any other
  raw OOXML `numFmt`), indentation, start-at value, per level - `NumberingDefinition.
  Levels` isn't limited to one level, and `Writer`/`Reader` handle any number of them the
  same way; `Builders.multiLevelNumberedListDef` builds the common correctly-linked
  outline shape ("1.", "1.1.", "1.1.1.", ...) without a caller having to hand-author each
  level's own `%N`-placeholder `Text` pattern (which this DSL still doesn't validate
  against the nesting it's used at - trust-the-caller, same posture direct formatting
  fields take). This DSL collapses WordprocessingML's own abstract-numbering/
  numbering-instance indirection into one `NumberingDefinition` per distinct list.
- **Tables**: rows and cells with horizontal merge (`GridSpan`) and vertical merge
  (`Restart`/`Continue`), per-cell shading and border overrides, column widths, a table
  style reference (name + banding flags - either a built-in like `"TableGrid"`, or a
  custom `TableStyleDefinition`, see below), table-level borders (outer sides plus inside
  horizontal/vertical), a table-wide default cell margin (`TableEntry.CellMargins`) with a
  per-cell override on top (`TableCellProps.Margins` - same `CellMargins` shape either
  way), and a row's own "repeat as header on every page" flag (`TableRow.
  RepeatAsHeader`). A cell's content is a `Block list`, so a cell containing a nested
  table is exactly how Word itself represents one.
- **Custom table style definitions** (`TableStyleDefinition`, `styles.xml` `w:type="table"`
  entries, referenced from `Document.TableStyles` by a `TableStyleRef.Name` the same way a
  built-in name is): base table borders, plus eleven of OOXML's thirteen conditional-
  formatting regions - whole-table defaults, first/last row, first/last column, one band
  per banding axis (`BandedRow`/`BandedColumn`), and all four corner cells - each with its
  own run/paragraph formatting and cell shading. See the gap below on the two regions not
  covered.
- **Images**: PNG/JPEG/GIF/BMP raster images anchored inline within a run (the natural
  placement for Word, unlike Excel's cell-range anchor) - `Data` is the image file's own
  raw bytes, embedded and read back byte-for-byte with no decoding/re-encoding.
- **Hyperlinks**: external (any URL, including `mailto:`) and internal (same-document
  bookmark) targets, wrapping one or more runs, with an optional tooltip.
- **Bookmarks**: named ranges within the document. `Inline.Bookmark` wraps inline content
  within a single paragraph, the common ergonomic case; a bookmark spanning more than one
  paragraph is `BookmarkRangeStart`/`BookmarkRangeEnd` instead - two independent markers a
  caller places directly in separate paragraphs' own `Inlines`, sharing a `name` (`Writer`
  assigns the matching OOXML `w:id` automatically; `Reader` resolves a bare `w:bookmarkEnd`
  back to its name via a document-wide id->name pass built before per-paragraph reading).
- **Comments**: modern Word comments (author, initials, date, text) anchored to a range of
  inline content. Not the legacy "cell comment" concept Excel models; this is what current
  Word's UI calls a comment. `Inline.Comment` wraps inline content within a single
  paragraph, the common ergonomic case; a comment spanning more than one paragraph is
  `CommentRangeStart`/`CommentRangeEnd` instead - two independent markers a caller places
  directly in separate paragraphs' own `Inlines`, sharing an `id`. Unlike `Bookmark`'s
  `name`, OOXML has nowhere to persist an arbitrary string alongside a comment range (only
  a numeric `w:id`, which `Writer` assigns itself) - `CommentRangeStart`'s own `id` is
  write-time-only, a correlation key for wiring the pair together when building the
  document; `Reader` reconstructs some id from the real OOXML `w:id` instead, which
  generally won't match what a caller originally wrote. Nothing else in a document ever
  references a comment by this id the way `InternalBookmark` can reference a bookmark's
  own `name`, so that's never a practical problem in use - it only means comparing a
  `CommentRangeStart`/`End`'s own `id` field byte-for-byte across a round trip isn't
  meaningful (`tests/Kookerella.FsWordDsl.Tests/Tests.fs`'s own round-trip assertion
  normalizes it away before comparing, the same treatment `DocumentProtection.Password`
  gets there for the same "known write-time-only field" reason).
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
  has - a caller never has to think about either. Numbering itself is customizable per
  section (`SectionProperties.FootnoteNumbering`/`EndnoteNumbering`, `w:sectPr/
  w:footnotePr`/`w:endnotePr`): number format (reusing `Numbering.NumberFormatKind`),
  start-at value, and restart behavior (continuous, or restarting each section/page) -
  `None` is Word's own default (continuous decimal from 1).
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
- **Document core properties**: Title, Author, Subject, Keywords, Comments, Category
  (`docProps/core.xml`, via `WordprocessingDocument.PackageProperties`) and Company
  (`docProps/app.xml`, via `ExtendedFilePropertiesPart`) - only written/read when at least
  one is set (`DocumentProperties.Default` round-trips to nothing on disk, same "all-defaults
  reads back as absent" discipline the rest of this DSL follows).
- **Track changes** (`w:ins`/`w:del`), narrowly scoped to the case that actually matters for
  almost every real redlined document: `Inline.TrackedChange` wraps arbitrary inline content
  the same way `Bookmark`/`Comment` do, marking it inserted or deleted with an author and a
  date; `Paragraph.MarkRevision` separately tracks whether the paragraph's own closing mark
  (not its content) was inserted or deleted. A run inside a `Deleted` `TrackedChange` writes
  its text as `w:delText` rather than `w:t` (schema-required) and reads back identically to
  an ordinary `Run` either way - the surrounding `TrackedChange` wrapper is what actually
  carries the deleted-ness, not the run itself. See the gap below on what's deliberately not
  covered (formatting-change history, moves, table row/cell tracking).

## Known gaps (documented, not silently "supported")

- **Track changes beyond inserted/deleted content and paragraph marks.** Formatting-change
  history (`w:rPrChange`/`w:pPrChange` - the run/paragraph formatting a revision *replaced*,
  for Word's own compare/undo view), moves (`w:moveFrom`/`w:moveTo` - Word's own "this looks
  like cut+paste" detection; without it, a moved block just round-trips as an ordinary
  delete-in-the-old-spot plus insert-in-the-new-spot, which is still correct information,
  just not the special annotation), and table row/cell-level insertion/deletion tracking
  (`w:trPr/w:ins`, `w:tcPr/w:cellIns`) aren't modeled - see "Modeled faithfully" above for
  what is.
- **Content controls (structured document tags), and `w:customXml` wrappers.** This DSL
  doesn't model the control itself (tag/title/lock/placeholder/data binding, dropdown/
  date/checkbox/... variants) - none of that is authored, and none of it survives a
  round trip. `Reader` does still recover the block/inline content a `w:sdt`/`w:customXml`
  *wraps* when reading a foreign file (unwrapping it rather than failing the whole
  document, since these are extremely common in real-world templates - see the note below
  on `Reader`'s own resilience posture).
- **Real field computation.** Only raw instruction text + a cached display value round-trip
  (see "Modeled faithfully" above) - a table of contents, cross-reference, or any other
  field that depends on document layout is never actually computed by this DSL, unlike
  Excel's pivot tables (which *do* perform real aggregation at write time).
- **The "second" band of each table-style banding axis.** `TableStyleDefinition.
  BandedRow`/`BandedColumn` apply to the odd/first band only (`w:type="band1Horz"`/
  `"band1Vert"`) - a distinct look for the even band (`band2Horz`/`band2Vert`) isn't
  modeled, since in practice a banded table's "off" band is just `WholeTable`'s own
  default background showing through (see "Modeled faithfully" above for the eleven
  regions that *are* covered).
- **Theme colors on borders, and theme colors used in `w:highlight`.** `BorderSide.Color`
  round-trips a `Theme` value as its `Fallback` RGB only (the theme token itself isn't
  preserved there - see "Modeled faithfully" above); `RunStyle.Highlight`'s fixed
  16-color palette has no theme-color concept in OOXML at all, so this doesn't apply to it.
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
