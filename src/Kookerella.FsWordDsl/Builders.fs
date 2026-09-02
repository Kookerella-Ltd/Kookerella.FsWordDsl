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
          VbaProject = None
          Properties = DocumentProperties.Default
          TableStyles = [] }

    /// Pipe-friendly, mirroring Excel's own `withDefinedNames`/`withProtection`.
    let withStyles (styles: StyleDefinition list) (doc: Document) : Document = { doc with Styles = styles }

    /// Title/Author/etc. - `Writer` only touches `docProps/core.xml`/`app.xml` at all when
    /// at least one field here is set, see `DocumentProperties`'s own doc comment.
    let withDocumentProperties (properties: DocumentProperties) (doc: Document) : Document = { doc with Properties = properties }

    let withNumbering (definitions: NumberingDefinition list) (doc: Document) : Document = { doc with Numbering = definitions }

    /// Pipe-friendly, mirroring `withStyles` - see `TableStyleDefinition`'s own doc comment.
    let withTableStyles (definitions: TableStyleDefinition list) (doc: Document) : Document = { doc with TableStyles = definitions }

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

    /// A multi-level decimal-numbered outline list ("1.", "1.1.", "1.1.1.", ...) -
    /// `Model.Paragraph.Numbering`'s own doc comment already notes `NumberingDefinition.
    /// Levels` supports several levels; this is just the common, correctly-linked shape
    /// (each level's `Text` chains through its ancestors' own counters via consecutive
    /// `%N` placeholders - hand-authoring that yourself is easy to get subtly wrong, since
    /// this DSL doesn't validate it, see `ListLevel.Text`'s own doc comment) with
    /// increasing indentation per level. `levelCount` must be between 1 and 9
    /// (WordprocessingML's own `w:lvl` range).
    let multiLevelNumberedListDef (id: int) (levelCount: int) : NumberingDefinition =
        { Id = id
          Levels =
            [ for i in 1..levelCount ->
                  { Format = DecimalFormat
                    Text = ([ 1..i ] |> List.map (sprintf "%%%d") |> String.concat ".") + "."
                    IndentLeft = Some(36.0 * float i)
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
            ?numbering: int * int,
            ?markRevision: Revision
        ) : Block =
        ParagraphBlock
            { Inlines = inlines
              StyleId = styleId
              Format = format
              Numbering = numbering
              MarkRevision = markRevision }

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
        Footnote [ ParagraphBlock { Inlines = [ Run(text, None, None) ]; StyleId = Some BuiltInStyles.footnoteTextStyle.Id; Format = None; Numbering = None; MarkRevision = None } ]

    /// A footnote wrapping already-built body content (several paragraphs, or a table).
    static member footnote(content: Block list) : Inline = Footnote content

    static member endnote(text: string) : Inline =
        Endnote [ ParagraphBlock { Inlines = [ Run(text, None, None) ]; StyleId = Some BuiltInStyles.endnoteTextStyle.Id; Format = None; Numbering = None; MarkRevision = None } ]

    static member endnote(content: Block list) : Inline = Endnote content

    /// Marks `content` as inserted under track changes (`w:ins`) - `date` defaults to
    /// "now" at write time when omitted, same convention `comment`'s own `date` uses.
    static member inserted(content: Inline list, author: string, ?date: DateTime) : Inline =
        TrackedChange({ Kind = Inserted; Author = author; Date = date }, content)

    /// Marks `content` as deleted under track changes (`w:del`).
    static member deleted(content: Inline list, author: string, ?date: DateTime) : Inline =
        TrackedChange({ Kind = Deleted; Author = author; Date = date }, content)

    /// A run-level content control (`w:sdt`), sitting inside a single paragraph. See
    /// `ContentControls.fs`'s own doc comment for what `controlType` can be.
    static member contentControl(content: Inline list, controlType: ContentControlType, ?alias: string, ?tag: string) : Inline =
        InlineContentControl({ Alias = alias; Tag = tag; Type = controlType }, content)

    /// The block-level counterpart, wrapping whole paragraphs/tables rather than sitting
    /// inside one paragraph.
    static member contentControlBlock(content: Block list, controlType: ContentControlType, ?alias: string, ?tag: string) : Block =
        ContentControlBlock({ Alias = alias; Tag = tag; Type = controlType }, content)

    static member tableCell(content: Block list, ?props: TableCellProps) : TableCell =
        { Content = content
          Props = defaultArg props TableCellProps.Default }

    static member tableRow(cells: TableCell list, ?height: float, ?repeatAsHeader: bool) : TableRow =
        { Cells = cells
          Height = height
          RepeatAsHeader = defaultArg repeatAsHeader false }

    static member table
        (
            rows: TableRow list,
            columnWidths: float list,
            ?style: TableStyleRef,
            ?borders: TableBorders,
            ?cellMargins: CellMargins
        ) : Block =
        TableBlock
            { Rows = rows
              ColumnWidths = columnWidths
              Style = style
              Borders = borders
              CellMargins = cellMargins }
