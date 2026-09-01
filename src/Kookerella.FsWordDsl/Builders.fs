namespace Kookerella.FsWordDsl

open System

/// Plain functions for constructing `Section`/`Document` values, plus `DocumentDsl` - real
/// optional-parameter smart constructors for `Inline`/`Block`/table pieces, the Word analog
/// of `SheetDsl`. Notably simpler than Excel's `SheetItem` fold here: paragraphs and table
/// rows are naturally sequential lists (no sparse row/column index to default from), so
/// there's no cursor-threading fold the way `SheetItems.sheet` needs - `para`/`table` build
/// their `Block` directly.
[<AutoOpen>]
module Builders =

    let section (body: Block list) : Section =
        { Body = body; Properties = SectionProperties.Default }

    let sectionWith (properties: SectionProperties) (body: Block list) : Section = { Body = body; Properties = properties }

    /// Builds a `Document` from one or more sections, defaulting `Styles` to
    /// `BuiltInStyles.all` so `StyleId = Some "Heading1"` (or any other built-in id) just
    /// works without the caller registering it first - pipe `withStyles` afterward to
    /// replace or extend that set.
    let document (sections: Section list) : Document =
        { Sections = sections
          Styles = BuiltInStyles.all
          Numbering = []
          Protection = None
          VbaProject = None }

    /// Pipe-friendly, mirroring Excel's own `withDefinedNames`/`withProtection`.
    let withStyles (styles: StyleDefinition list) (doc: Document) : Document = { doc with Styles = styles }

    let withNumbering (definitions: NumberingDefinition list) (doc: Document) : Document = { doc with Numbering = definitions }

    let withProtection (protection: DocumentProtection) (doc: Document) : Document = { doc with Protection = Some protection }

    /// See `Document.VbaProject`'s own doc comment for what this does and doesn't cover.
    let withVbaProject (vbaProjectBytes: byte[]) (doc: Document) : Document = { doc with VbaProject = Some vbaProjectBytes }

    /// A single-level bullet list definition using Word's own conventional bullet glyph
    /// (rendered from the Symbol font, matching a fresh "Bullets" list in real Word).
    let bulletListDef (id: int) : NumberingDefinition =
        { Id = id
          Levels =
            [ { Format = BulletFormat(char 0xF0B7, "Symbol")
                Text = string (char 0xF0B7)
                IndentLeft = Some 36.0
                HangingIndent = Some 18.0
                StartAt = None } ] }

    /// A single-level decimal-numbered list definition ("1.", "2.", "3.", ...).
    let numberedListDef (id: int) : NumberingDefinition =
        { Id = id
          Levels =
            [ { Format = DecimalFormat
                Text = "%1."
                IndentLeft = Some 36.0
                HangingIndent = Some 18.0
                StartAt = Some 1 } ] }

/// Smart constructors, as members with real optional parameters - plain `let` bindings
/// can't have optional parameters in F# (member-only), same reason `SheetDsl` exists.
/// `open type Kookerella.FsWordDsl.DocumentDsl` (alongside `open Kookerella.FsWordDsl`)
/// brings `run`/`para`/... into scope unqualified.
type DocumentDsl =

    /// `styleId` references a character style (e.g. `"Hyperlink"`); `style` is direct
    /// formatting layered on top - either, both, or neither may be given.
    static member run(text: string, ?style: RunStyle, ?styleId: string) : Inline = Run(text, style, styleId)

    static member para
        (
            inlines: Inline list,
            ?styleId: string,
            ?format: ParagraphFormat,
            ?numbering: int * int
        ) : Block =
        ParagraphBlock
            { Inlines = inlines
              StyleId = styleId
              Format = format
              Numbering = numbering }

    /// A hyperlink over plain text - applies `BuiltInStyles.hyperlinkCharStyle` (blue,
    /// underlined) automatically so callers don't have to restate it on every run.
    static member hyperlink(text: string, target: HyperlinkTarget, ?tooltip: string) : Inline =
        Hyperlink(target, [ Run(text, None, Some BuiltInStyles.hyperlinkCharStyle.Id) ], tooltip)

    /// A hyperlink wrapping already-built runs, for mixed formatting within the link text.
    static member hyperlink(runs: Inline list, target: HyperlinkTarget, ?tooltip: string) : Inline =
        Hyperlink(target, runs, tooltip)

    static member bookmark(name: string, content: Inline list) : Inline = Bookmark(name, content)

    /// `author` defaults to an empty (unnamed) author, matching Excel's own `SheetDsl.
    /// comment`; `date` defaults to "now" at write time when omitted.
    static member comment
        (
            content: Inline list,
            text: string,
            ?author: string,
            ?initials: string,
            ?date: DateTime
        ) : Inline =
        Comment(defaultArg author "", initials, date, text, content)

    static member image(entry: ImageEntry) : Inline = Image entry

    /// A footnote over plain text - applies `BuiltInStyles.footnoteTextStyle` to the note
    /// body's own paragraph, same "caller doesn't restate the built-in id" convenience
    /// `hyperlink`'s text overload gives.
    static member footnote(text: string) : Inline =
        Footnote [ ParagraphBlock { Inlines = [ Run(text, None, None) ]; StyleId = Some BuiltInStyles.footnoteTextStyle.Id; Format = None; Numbering = None } ]

    /// A footnote wrapping already-built body content (several paragraphs, or a table).
    static member footnote(content: Block list) : Inline = Footnote content

    static member endnote(text: string) : Inline =
        Endnote [ ParagraphBlock { Inlines = [ Run(text, None, None) ]; StyleId = Some BuiltInStyles.endnoteTextStyle.Id; Format = None; Numbering = None } ]

    static member endnote(content: Block list) : Inline = Endnote content

    static member tableCell(content: Block list, ?props: TableCellProps) : TableCell =
        { Content = content
          Props = defaultArg props TableCellProps.Default }

    static member tableRow(cells: TableCell list, ?height: float) : TableRow = { Cells = cells; Height = height }

    static member table
        (
            rows: TableRow list,
            columnWidths: float list,
            ?style: TableStyleRef,
            ?borders: TableBorders
        ) : Block =
        TableBlock
            { Rows = rows
              ColumnWidths = columnWidths
              Style = style
              Borders = borders }
