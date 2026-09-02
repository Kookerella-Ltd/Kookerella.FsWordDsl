namespace Kookerella.FsWordDsl

open System

/// The document content model. Word's own structure is genuinely recursive - a table
/// cell's content is itself a `Block list` that can contain another table, and a header/
/// footer's content is a `Block list` too - so `Block`, `TableCell`/`TableRow`/`TableEntry`,
/// `HeaderFooterSet`/`SectionProperties`, and `Section` are declared together here via a
/// mutually-recursive `and` chain, rather than split one-type-per-file the way Excel's
/// non-recursive model allows (see `Tables.fs`/`PageSetup.fs`'s own notes on why their
/// content-holding types moved here instead).
///
/// Several case/type names below deliberately match `DocumentFormat.OpenXml.Wordprocessing`
/// types of the same name (`Paragraph`... no, see `Block`'s own note; `Table`/`TableRow`/
/// `TableCell`, `Hyperlink`, `Bookmark`, `Comment`) - `Interpreter/Writer.fs`/`Reader.fs`
/// never `open` that namespace directly (F# can't alias a namespace as a module), they
/// `open DocumentFormat.OpenXml` and always qualify via the nested namespace's own short
/// name (`Wordprocessing.Table(...)`), so these natural names stay available unqualified in
/// the DSL itself with no ambiguity at either call site.
[<AutoOpen>]
module Model =

    /// A paragraph's inline content. `Run` is Word's own paragraph-level content unit -
    /// unlike Excel's `CellValue` (always one uniformly-styled `Text` cell), a `Paragraph`'s
    /// `Inlines` naturally hold several independently-styled runs, so rich text (mixed
    /// formatting within one paragraph) is first-class here, not a documented gap.
    type Inline =
        /// `styleId` references a character style (`Document.Styles`, e.g. `"Hyperlink"`)
        /// - the same StyleId/direct-formatting split `Paragraph` has, one level down.
        | Run of text: string * style: RunStyle option * styleId: string option
        | LineBreak
        | Tab
        /// An explicit manual page break, same as pressing Ctrl+Enter mid-paragraph -
        /// distinct from `ParagraphFormat.PageBreakBefore`, which breaks before an entire
        /// paragraph rather than mid-run.
        | PageBreak
        | Image of ImageEntry
        | Hyperlink of target: HyperlinkTarget * runs: Inline list * tooltip: string option
        /// Wraps its own content, scoped to within a single paragraph - the ergonomic,
        /// common case. A bookmark spanning more than one paragraph is `BookmarkRangeStart`/
        /// `BookmarkRangeEnd` instead - two independent markers a caller places directly in
        /// separate `Paragraph.Inlines` lists (see those cases' own doc comments).
        | Bookmark of name: string * content: Inline list
        /// A bookmark boundary marker usable on its own, for a bookmark spanning more than
        /// one paragraph - `name` must match its corresponding `BookmarkRangeEnd`
        /// elsewhere in the same document (`Writer` assigns the matching OOXML `w:id`
        /// automatically, same as `Bookmark` does). Order matters: `Reader` treats a
        /// `BookmarkRangeStart` with no matching `BookmarkRangeEnd` later in the body as
        /// unterminated and drops it, same "best-effort on a foreign file" posture the rest
        /// of `Reader` takes.
        | BookmarkRangeStart of name: string
        | BookmarkRangeEnd of name: string
        /// `Date = None` is written as "now" at write time - Word records a comment's own
        /// timestamp, this DSL doesn't require the caller to supply one. Wraps its own
        /// content, scoped to within a single paragraph - the ergonomic, common case. A
        /// comment spanning more than one paragraph is `CommentRangeStart`/`CommentRangeEnd`
        /// instead (see those cases' own doc comments).
        | Comment of author: string * initials: string option * date: DateTime option * text: string * content: Inline list
        /// A comment boundary marker usable on its own, for a comment spanning more than one
        /// paragraph - unlike `Comment`, which wraps its own content, this carries the
        /// comment's own metadata itself (there's no wrapping case here to hang it off).
        /// `id` is a caller-chosen correlation key matching this start to its
        /// `CommentRangeEnd` elsewhere in the document - unlike `Bookmark`'s `name`, OOXML
        /// has nowhere to persist an arbitrary string alongside a comment range (only a
        /// numeric `w:id`, which `Writer` assigns itself), so `id` is write-time-only: it
        /// does not round-trip, `Reader` reconstructs some id from the real OOXML `w:id`
        /// instead, which will generally not match what you wrote. Nothing else in a
        /// document ever references a comment by this id the way `InternalBookmark` can
        /// reference a bookmark's own `name`, so that's never a practical problem.
        | CommentRangeStart of id: string * author: string * initials: string option * date: DateTime option * text: string
        /// Closes the range opened by the `CommentRangeStart` with the same `id` earlier in
        /// the document body. Order matters: `Reader` treats a `CommentRangeStart` with no
        /// matching `CommentRangeEnd` later in the body as unterminated and drops it, same
        /// "best-effort on a foreign file" posture the rest of `Reader` takes.
        | CommentRangeEnd of id: string
        /// A "simple field" (`w:fldSimple`) - raw field instruction text (e.g. `"PAGE"`,
        /// `"DATE \\@ \"MMMM d, yyyy\""`) plus the cached display text Word showed before
        /// its own recalculation. This DSL never evaluates a field itself, the same
        /// "cachedValue is the only number that will ever exist until something else
        /// computes one" posture Excel's `CellValue.Formula` documents - real Word
        /// recalculates on open and overwrites it.
        | Field of instruction: string * cachedResult: string option
        /// A footnote/endnote reference mark, e.g. `Footnote [ ParagraphBlock { ... } ]` -
        /// `content` is the note's own body (own paragraphs, possibly a table), stored in
        /// `word/footnotes.xml`/`endnotes.xml` and referenced from here by an id `Writer`
        /// assigns automatically. `Writer` also prepends the note-reference-mark run
        /// (`w:footnoteRef`/`w:endnoteRef`) to the body's first paragraph itself - a caller
        /// writes ordinary paragraph content and never has to think about that marker.
        | Footnote of content: Block list
        | Endnote of content: Block list

    /// One paragraph's content plus its formatting. `StyleId` references a named style
    /// (`Document.Styles`, e.g. `"Heading1"`) - the inheritance layer; `Format` is direct/
    /// inline formatting layered on top, same relationship `Styles.ParagraphFormat`'s own
    /// doc comment describes. `Numbering = Some(numId, level)` places this paragraph in a
    /// numbered/bulleted list (see `Numbering.fs`).
    and Paragraph =
        { Inlines: Inline list
          StyleId: string option
          Format: ParagraphFormat option
          Numbering: (int * int) option }

    /// A block-level unit of document content - a paragraph, or a table (which itself
    /// contains more `Block`s per cell, hence the recursion). Unlike `Inline`'s cases,
    /// `ParagraphBlock`/`TableBlock` are DSL-only names with no single corresponding OOXML
    /// element - real WordprocessingML body content is just a flat sequence of `<w:p>`/
    /// `<w:tbl>` elements, `Block` exists so this DSL can describe that sequence as one list.
    /// Joins the same recursive `and` chain as `TableCell` etc. below now that `Inline.
    /// Footnote`/`Endnote` close a cycle back through here (a note's own body is itself a
    /// `Block list`).
    and Block =
        | ParagraphBlock of Paragraph
        | TableBlock of TableEntry

    /// One table cell. `Content` is almost always a single `ParagraphBlock` in practice
    /// (Word requires at least one paragraph per cell even when it's empty), but nothing
    /// here enforces that - a cell containing a nested `TableBlock` is exactly how Word
    /// itself represents a nested table.
    and TableCell = { Content: Block list; Props: TableCellProps }

    and TableRow =
        { Cells: TableCell list
          /// Points.
          Height: float option
          /// Repeats this row at the top of every page the table spans (`w:trPr/
          /// w:tblHeader`) - meaningful only on a table's leading row(s); real Word ignores
          /// it on a row that isn't part of a contiguous run starting at row 1.
          RepeatAsHeader: bool }

    /// `ColumnWidths` gives the table's own grid (`w:tblGrid`) - one entry per column, in
    /// points; a row's own cells should sum to the same column count accounting for any
    /// `GridSpan`, the same shape/width validation Excel's own `Table` performs against its
    /// range width (`Interpreter/Writer.fs` validates this the same way at write time).
    and TableEntry =
        { Rows: TableRow list
          ColumnWidths: float list
          Style: TableStyleRef option
          Borders: TableBorders option
          /// The table's own default cell margins (`w:tblCellMar`) - see `CellMargins`'s own
          /// doc comment for what this does and doesn't cover.
          CellMargins: CellMargins option }

    /// `Default`/`First`/`Even` mirror Word's own three header/footer variants exactly -
    /// `First` shows only on a section's first page (requires the sibling
    /// `titlePg`/`differentFirst` flag, which `Writer` sets automatically whenever this is
    /// `Some`, same auto-flag convention Excel's own `PageSetup.FirstHeader` uses), `Even`
    /// only on even pages (requires `differentOddEven`/`evenAndOddHeaders`, same auto-flag
    /// treatment). `Default` covers odd pages when `Even` is set, or every page otherwise.
    and HeaderFooterSet =
        { Default: Block list option
          First: Block list option
          Even: Block list option }

        static member None =
            { Default = None
              First = None
              Even = None }

    /// One section's page setup - a document is a sequence of `Section`s (see `Document`
    /// below), each with its own, mapping 1:1 onto a real Word section break rather than
    /// needing a synthetic "section break" block type. `BreakType` is how *this* section
    /// begins relative to the previous one (see `PageSetup.SectionBreakType`'s own doc
    /// comment) - meaningless, and not written, for the very first section, since there's
    /// no previous section for it to break from. `Columns` is the number of equal-width
    /// text columns (1 = the ordinary single-column case).
    and SectionProperties =
        { PageSize: PageSize
          Orientation: PageOrientation
          Margins: PageMargins
          Header: HeaderFooterSet option
          Footer: HeaderFooterSet option
          PageNumberStart: int option
          Columns: int
          BreakType: SectionBreakType
          /// `None` is Word's own default (continuous decimal numbering starting at 1,
          /// straight through the document) - see `PageSetup.NoteNumberingSettings`'s own
          /// doc comment.
          FootnoteNumbering: NoteNumberingSettings option
          EndnoteNumbering: NoteNumberingSettings option }

        static member Default =
            { PageSize = Letter
              Orientation = Portrait
              Margins = PageMargins.Default
              Header = None
              Footer = None
              PageNumberStart = None
              Columns = 1
              BreakType = NextPageBreak
              FootnoteNumbering = None
              EndnoteNumbering = None }

    and Section = { Body: Block list; Properties: SectionProperties }

    /// A macro-enabled template's VBA project, stored as its own file's raw bytes exactly
    /// as they'd sit in `word/vbaProject.bin` - a compiled OLE/CFBF binary blob, not source
    /// text, given the exact same "opaque payload, no authoring" treatment Excel's own
    /// `Workbook.VbaProject` documents (see that type's own doc comment for the full
    /// reasoning, which applies here unchanged). Presence of a VBA project switches the
    /// saved file's content type to Word's macro-enabled kind (`Writer` picks
    /// `WordprocessingDocumentType.MacroEnabledDocument`/`.MacroEnabledTemplate` as
    /// appropriate); save with a `.docm`/`.dotm` path for real Word to actually trust and
    /// run it.
    type Document =
        { Sections: Section list
          Styles: StyleDefinition list
          Numbering: NumberingDefinition list
          Protection: DocumentProtection option
          VbaProject: byte[] option
          /// Title/Author/Subject/etc. - see `DocumentProperties`'s own doc comment.
          Properties: DocumentProperties
          /// Custom table style *definitions* - see `TableStyleDefinition`'s own doc
          /// comment. Referenced from a `TableEntry.Style.Name` the same way a built-in
          /// name like `"TableGrid"` is.
          TableStyles: TableStyleDefinition list }
