namespace Kookerella.FsWordDsl.Interpreter

open System
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open Kookerella.FsWordDsl
open Kookerella.FsWordDsl.Interpreter.StyleRegistry
open Kookerella.FsWordDsl.Interpreter.ImageReader

/// OOXML -> DSL, the reverse transform. Round-trip correctness is guaranteed and tested
/// against `Writer`'s own output (see this repo's `Tests.fs`) - reading an arbitrary
/// real-world `.docx` is best-effort, same "drop what isn't modeled" posture Excel's own
/// `Reader` takes (see MAPPING.md).
module Reader =

    let private opt (x: 'a when 'a: null) : 'a option = if obj.ReferenceEquals(x, null) then None else Some x

    /// Shared read-time lookups, threaded through the same functions that used to thread a
    /// bare `commentsById` map alone - `FootnotesById`/`EndnotesById` key a note's already-
    /// parsed `Block list` body by its `w:id`, built once in `readDocument` before the body
    /// itself is read (see that function's own note on why note bodies get an empty-notes
    /// bootstrap `rctx` rather than needing this to be lazy/two-pass).
    type private ReadCtx =
        { CommentsById: Map<string, Wordprocessing.Comment>
          FootnotesById: Map<int, Block list>
          EndnotesById: Map<int, Block list>
          /// Every `w:bookmarkStart`'s own id -> name, built once from the whole document
          /// body before per-paragraph reading begins - needed because a `BookmarkEnd`
          /// reached on its own (no matching `BookmarkStart` earlier in the same
          /// paragraph's own `elements` array, i.e. the other half of a cross-paragraph
          /// bookmark) carries only an id, never a name.
          BookmarkNamesById: Map<string, string> }

    // --- Numbering --------------------------------------------------------------------------

    let private numberFormatKindOfW (v: Wordprocessing.NumberFormatValues) : NumberFormatKind =
        if v = Wordprocessing.NumberFormatValues.Bullet then BulletFormat(char 0x2022, "Symbol") // glyph/font overridden below from the level's own text/rPr
        elif v = Wordprocessing.NumberFormatValues.Decimal then DecimalFormat
        elif v = Wordprocessing.NumberFormatValues.LowerLetter then LowerLetterFormat
        elif v = Wordprocessing.NumberFormatValues.UpperLetter then UpperLetterFormat
        elif v = Wordprocessing.NumberFormatValues.LowerRoman then LowerRomanFormat
        elif v = Wordprocessing.NumberFormatValues.UpperRoman then UpperRomanFormat
        else OtherFormat(v.ToString())

    let private twipsToPoints (s: string) : float = float (int s) / 20.0

    let private levelOfW (lvl: Wordprocessing.Level) : ListLevel =
        let baseFormat =
            lvl.NumberingFormat
            |> opt
            |> Option.bind (fun nf -> nf.Val |> opt)
            |> Option.map (fun v -> numberFormatKindOfW v.Value)
            |> Option.defaultValue DecimalFormat

        let text = lvl.LevelText |> opt |> Option.bind (fun t -> t.Val |> opt) |> Option.map (fun v -> v.Value) |> Option.defaultValue "%1."

        let format =
            match baseFormat with
            | BulletFormat(defaultGlyph, _) ->
                let font =
                    lvl.NumberingSymbolRunProperties
                    |> opt
                    |> Option.bind (fun rpr -> rpr.RunFonts |> opt)
                    |> Option.bind (fun rf -> rf.Ascii |> opt)
                    |> Option.map (fun v -> v.Value)
                    |> Option.defaultValue "Symbol"

                // The bullet's actual glyph is its own level text (a single character) -
                // `NumberFormatValues.Bullet` alone doesn't carry which glyph, only that
                // this level is a bullet at all.
                let glyph = if text.Length > 0 then text.[0] else defaultGlyph
                BulletFormat(glyph, font)
            | other -> other

        let indentLeft, hanging =
            match lvl.PreviousParagraphProperties |> opt |> Option.bind (fun p -> p.Indentation |> opt) with
            | None -> None, None
            | Some ind ->
                (ind.Left |> opt |> Option.map (fun v -> twipsToPoints v.Value)), (ind.Hanging |> opt |> Option.map (fun v -> twipsToPoints v.Value))

        let startAt =
            lvl.StartNumberingValue |> opt |> Option.bind (fun s -> s.Val |> opt) |> Option.map (fun v -> v.Value)

        { Format = format
          Text = text
          IndentLeft = indentLeft
          HangingIndent = hanging
          StartAt = startAt }

    let private numberingOfW (numbering: Wordprocessing.Numbering option) : NumberingDefinition list =
        match numbering with
        | None -> []
        | Some numbering ->
            let abstractLevels =
                numbering.Elements<Wordprocessing.AbstractNum>()
                |> Seq.map (fun absNum ->
                    let id = absNum.AbstractNumberId.Value

                    let levels =
                        absNum.Elements<Wordprocessing.Level>()
                        |> Seq.sortBy (fun l -> l.LevelIndex.Value)
                        |> Seq.map levelOfW
                        |> List.ofSeq

                    id, levels)
                |> Map.ofSeq

            numbering.Elements<Wordprocessing.NumberingInstance>()
            |> Seq.choose (fun inst ->
                let numId = inst.NumberID.Value
                let abstractId = inst.Elements<Wordprocessing.AbstractNumId>() |> Seq.tryHead |> Option.map (fun a -> a.Val.Value)

                abstractId
                |> Option.bind (fun aid -> abstractLevels.TryFind aid)
                |> Option.map (fun levels -> { Id = numId; Levels = levels }))
            |> List.ofSeq

    // --- Inline content ---------------------------------------------------------------------

    let private readRunInlines (r: Wordprocessing.Run) : Inline list =
        if r.Elements<Wordprocessing.CommentReference>() |> Seq.isEmpty |> not then
            []
        else
            let style = runStyleOfProperties (r.RunProperties |> opt)
            let styleId = styleIdOfProperties (r.RunProperties |> opt)

            r.ChildElements
            |> Seq.choose (fun c ->
                match c with
                | :? Wordprocessing.Text as t -> Some(Run(t.Text, style, styleId))
                | :? Wordprocessing.Break as b ->
                    if b.Type |> opt |> Option.map (fun v -> v.Value) = Some Wordprocessing.BreakValues.Page then
                        Some PageBreak
                    else
                        Some LineBreak
                | :? Wordprocessing.TabChar -> Some Tab
                | _ -> None)
            |> List.ofSeq

    /// Looks up a comment's own metadata by its real OOXML id - shared by the wrapping
    /// `Comment` case and `CommentRangeStart` below, which differ only in how the
    /// surrounding range/content is read.
    let private commentMetaOfId (rctx: ReadCtx) (id: string) : string * string option * DateTime option * string =
        let meta = rctx.CommentsById.TryFind id
        let author = meta |> Option.map (fun c -> c.Author.Value) |> Option.defaultValue ""
        let initials = meta |> Option.bind (fun c -> c.Initials |> opt) |> Option.map (fun v -> v.Value)
        let date = meta |> Option.bind (fun c -> c.Date |> opt) |> Option.map (fun v -> v.Value)
        let text = meta |> Option.map (fun c -> c.InnerText) |> Option.defaultValue ""
        author, initials, date, text

    let rec private parseInlineRange
        (elements: OpenXmlElement[])
        (startIdx: int)
        (endIdxExclusive: int)
        (mainPart: MainDocumentPart)
        (rctx: ReadCtx)
        : Inline list =
        let result = ResizeArray<Inline>()
        let mutable i = startIdx

        while i < endIdxExclusive do
            match elements.[i] with
            | :? Wordprocessing.BookmarkStart as bs ->
                let id = bs.Id.Value
                let name = bs.Name.Value

                let endIdx =
                    seq { i + 1 .. endIdxExclusive - 1 }
                    |> Seq.tryFind (fun j ->
                        match elements.[j] with
                        | :? Wordprocessing.BookmarkEnd as be -> be.Id.Value = id
                        | _ -> false)

                match endIdx with
                | Some endIdx ->
                    let content = parseInlineRange elements (i + 1) endIdx mainPart rctx
                    result.Add(Bookmark(name, content))
                    i <- endIdx + 1
                | None ->
                    // No matching `BookmarkEnd` within this same paragraph's own elements -
                    // its other half is in a later paragraph, so this is one end of a
                    // cross-paragraph bookmark rather than the single-paragraph `Bookmark`
                    // case.
                    result.Add(BookmarkRangeStart name)
                    i <- i + 1
            | :? Wordprocessing.BookmarkEnd as be ->
                // Reached directly (not consumed by the `BookmarkStart` case above), so its
                // `BookmarkStart` was in an earlier paragraph - the closing half of a
                // cross-paragraph bookmark. `BookmarkNamesById` supplies the name, which a
                // bare `w:bookmarkEnd` never carries itself.
                match rctx.BookmarkNamesById.TryFind be.Id.Value with
                | Some name -> result.Add(BookmarkRangeEnd name)
                | None -> ()

                i <- i + 1
            | :? Wordprocessing.CommentRangeStart as crs ->
                let id = crs.Id.Value

                let endIdx =
                    seq { i + 1 .. endIdxExclusive - 1 }
                    |> Seq.tryFind (fun j ->
                        match elements.[j] with
                        | :? Wordprocessing.CommentRangeEnd as cre -> cre.Id.Value = id
                        | _ -> false)

                match endIdx with
                | Some endIdx ->
                    let content = parseInlineRange elements (i + 1) endIdx mainPart rctx
                    let afterEnd = endIdx + 1

                    let refConsumed =
                        afterEnd < endIdxExclusive
                        && (match elements.[afterEnd] with
                            | :? Wordprocessing.Run as r -> r.Elements<Wordprocessing.CommentReference>() |> Seq.exists (fun cr -> cr.Id.Value = id)
                            | _ -> false)

                    let author, initials, date, text = commentMetaOfId rctx id
                    result.Add(Comment(author, initials, date, text, content))
                    i <- if refConsumed then afterEnd + 1 else afterEnd
                | None ->
                    // No matching `CommentRangeEnd` within this same paragraph's own
                    // elements - its other half is in a later paragraph, so this is one end
                    // of a cross-paragraph comment rather than the single-paragraph
                    // `Comment` case. The reconstructed `id` is the real OOXML id, not
                    // whatever string a caller originally wrote (see `CommentRangeStart`'s
                    // own doc comment on why that's fine).
                    let author, initials, date, text = commentMetaOfId rctx id
                    result.Add(CommentRangeStart(id, author, initials, date, text))
                    i <- i + 1
            | :? Wordprocessing.CommentRangeEnd as cre ->
                // Reached directly (not consumed by the `CommentRangeStart` case above), so
                // its `CommentRangeStart` was in an earlier paragraph - the closing half of
                // a cross-paragraph comment.
                let id = cre.Id.Value
                result.Add(CommentRangeEnd id)
                let afterEnd = i + 1

                let refConsumed =
                    afterEnd < endIdxExclusive
                    && (match elements.[afterEnd] with
                        | :? Wordprocessing.Run as r -> r.Elements<Wordprocessing.CommentReference>() |> Seq.exists (fun cr -> cr.Id.Value = id)
                        | _ -> false)

                i <- if refConsumed then afterEnd + 1 else afterEnd
            | :? Wordprocessing.Hyperlink as hl ->
                let children = hl.ChildElements |> Seq.cast<OpenXmlElement> |> Array.ofSeq
                let innerInlines = parseInlineRange children 0 children.Length mainPart rctx
                let tooltip = hl.Tooltip |> opt |> Option.map (fun v -> v.Value)

                let target =
                    match hl.Anchor |> opt with
                    | Some anchor -> InternalBookmark anchor.Value
                    | None ->
                        match hl.Id |> opt with
                        | Some relId ->
                            match mainPart.HyperlinkRelationships |> Seq.tryFind (fun r -> r.Id = relId.Value) with
                            | Some rel -> ExternalUrl(rel.Uri.ToString())
                            | None -> ExternalUrl ""
                        | None -> ExternalUrl ""

                result.Add(Hyperlink(target, innerInlines, tooltip))
                i <- i + 1
            | :? Wordprocessing.Run as r ->
                let footnoteRef = r.Elements<Wordprocessing.FootnoteReference>() |> Seq.tryHead
                let endnoteRef = r.Elements<Wordprocessing.EndnoteReference>() |> Seq.tryHead

                match footnoteRef, endnoteRef with
                | Some fr, _ -> rctx.FootnotesById.TryFind(int fr.Id.Value) |> Option.iter (fun content -> result.Add(Footnote content))
                | _, Some er -> rctx.EndnotesById.TryFind(int er.Id.Value) |> Option.iter (fun content -> result.Add(Endnote content))
                | None, None ->
                    result.AddRange(readRunInlines r)

                    match r.Descendants<Wordprocessing.Drawing>() |> Seq.tryHead with
                    | Some drawing -> tryReadImage mainPart drawing |> Option.iter (fun img -> result.Add(Image img))
                    | None -> ()

                i <- i + 1
            | :? Wordprocessing.SimpleField as sf ->
                let instr = sf.Instruction |> opt |> Option.map (fun v -> v.Value) |> Option.defaultValue ""
                let cached = sf.Descendants<Wordprocessing.Text>() |> Seq.tryHead |> Option.map (fun t -> t.Text)
                result.Add(Field(instr, cached))
                i <- i + 1
            | _ -> i <- i + 1

        result |> List.ofSeq

    // --- Paragraphs / tables -----------------------------------------------------------------

    let private readParagraph (mainPart: MainDocumentPart) (rctx: ReadCtx) (p: Wordprocessing.Paragraph) : Paragraph =
        let pPr = p.ParagraphProperties |> opt
        let styleId = styleIdOfParagraphProperties pPr
        let format = paragraphFormatOfProperties pPr

        let numbering =
            pPr
            |> Option.bind (fun pr -> pr.NumberingProperties |> opt)
            |> Option.bind (fun np ->
                match np.NumberingId |> opt, np.NumberingLevelReference |> opt with
                | Some numId, Some lvl -> Some(numId.Val.Value, lvl.Val.Value)
                | _ -> None)

        let elements =
            p.ChildElements
            |> Seq.filter (fun c -> not (c :? Wordprocessing.ParagraphProperties))
            |> Array.ofSeq

        { Inlines = parseInlineRange elements 0 elements.Length mainPart rctx
          StyleId = styleId
          Format = format
          Numbering = numbering }

    let private tableStyleRefOfW (tblPr: Wordprocessing.TableProperties option) : TableStyleRef option =
        tblPr
        |> Option.bind (fun p -> p.TableStyle |> opt)
        |> Option.map (fun s ->
            let look = tblPr |> Option.bind (fun p -> p.TableLook |> opt)
            let flag (f: Wordprocessing.TableLook -> OnOffValue) = look |> Option.bind (fun l -> f l |> opt) |> Option.map (fun v -> v.Value) |> Option.defaultValue false

            { Name = s.Val.Value
              FirstRowBanding = flag (fun l -> l.FirstRow)
              LastRowBanding = flag (fun l -> l.LastRow)
              BandedRows = not (flag (fun l -> l.NoHorizontalBand))
              BandedColumns = not (flag (fun l -> l.NoVerticalBand)) })

    // `borderSideOfTop`/`OfBottom`/`OfLeft`/`OfRight`/`OfInsideH`/`OfInsideV` live in
    // `StyleRegistry.fs` now - shared with paragraph borders (`w:pBdr`).

    let private tableBordersOfW (tblPr: Wordprocessing.TableProperties option) : TableBorders option =
        tblPr
        |> Option.bind (fun p -> p.TableBorders |> opt)
        |> Option.map (fun tb ->
            { Outer =
                { Left = tb.LeftBorder |> opt |> Option.map borderSideOfLeft
                  Right = tb.RightBorder |> opt |> Option.map borderSideOfRight
                  Top = tb.TopBorder |> opt |> Option.map borderSideOfTop
                  Bottom = tb.BottomBorder |> opt |> Option.map borderSideOfBottom }
              InsideHorizontal = tb.InsideHorizontalBorder |> opt |> Option.map borderSideOfInsideH
              InsideVertical = tb.InsideVerticalBorder |> opt |> Option.map borderSideOfInsideV })

    /// Zero-or-more, not exactly one: an unrecognized body/cell-level element is dropped
    /// (`[]`) rather than failing the whole document, and a content control/custom XML
    /// wrapper expands to whatever recognized content it wraps. Reading an arbitrary
    /// real-world `.docx` is meant to be best-effort (see this module's own doc comment) -
    /// this used to be a lie for anything at block level it didn't recognize, throwing
    /// instead of degrading, which a `w:sdt` (a content control - extremely common in
    /// real-world templates: form fields, dropdowns, date pickers) hit constantly.
    let rec private readBlock (mainPart: MainDocumentPart) (rctx: ReadCtx) (el: OpenXmlElement) : Block list =
        match el with
        | :? Wordprocessing.Paragraph as p -> [ ParagraphBlock(readParagraph mainPart rctx p) ]
        | :? Wordprocessing.Table as t -> [ TableBlock(readTable mainPart rctx t) ]
        | :? Wordprocessing.SdtBlock as sdt ->
            // This DSL doesn't model content controls themselves (see MAPPING.md) - only
            // the block content they wrap is recovered, the control wrapper itself
            // (tag/title/lock/placeholder/data binding, ...) is silently discarded.
            match sdt.SdtContentBlock |> opt with
            | Some content -> content.ChildElements |> Seq.collect (readBlock mainPart rctx) |> List.ofSeq
            | None -> []
        | :? Wordprocessing.CustomXmlBlock as cx -> cx.ChildElements |> Seq.collect (readBlock mainPart rctx) |> List.ofSeq
        | _ ->
            // Anything else unmodeled at block level (`w:altChunk` - an embedded foreign
            // document format this DSL has no way to parse; a bookmark/comment range
            // marker placed directly in the body rather than nested in a paragraph;
            // `w:permStart`/`permEnd`; ...) - dropped, same "best-effort on a foreign
            // file" posture the rest of `Reader` takes.
            []

    and private readTable (mainPart: MainDocumentPart) (rctx: ReadCtx) (t: Wordprocessing.Table) : TableEntry =
        let tblPr = t.GetFirstChild<Wordprocessing.TableProperties>() |> opt
        let grid = t.GetFirstChild<Wordprocessing.TableGrid>() |> opt

        let widths =
            match grid with
            | None -> []
            | Some g -> g.Elements<Wordprocessing.GridColumn>() |> Seq.map (fun c -> twipsToPoints c.Width.Value) |> List.ofSeq

        let rows =
            t.Elements<Wordprocessing.TableRow>()
            |> Seq.map (fun tr ->
                let trPr = tr.GetFirstChild<Wordprocessing.TableRowProperties>() |> opt

                let height =
                    trPr
                    |> Option.bind (fun p -> p.GetFirstChild<Wordprocessing.TableRowHeight>() |> opt)
                    |> Option.bind (fun h -> h.Val |> opt)
                    |> Option.map (fun v -> twipsToPoints (string v.Value))

                let repeatAsHeader =
                    trPr |> Option.bind (fun p -> p.GetFirstChild<Wordprocessing.TableHeader>() |> opt) |> Option.isSome

                let cells =
                    tr.Elements<Wordprocessing.TableCell>()
                    |> Seq.map (fun tc ->
                        let tcPr = tc.TableCellProperties |> opt

                        let props =
                            { GridSpan = tcPr |> Option.bind (fun p -> p.GridSpan |> opt) |> Option.bind (fun g -> g.Val |> opt) |> Option.map (fun v -> v.Value)
                              VerticalMerge =
                                tcPr
                                |> Option.bind (fun p -> p.VerticalMerge |> opt)
                                |> Option.map (fun vm ->
                                    match vm.Val |> opt with
                                    | Some v when v.Value = Wordprocessing.MergedCellValues.Continue -> ContinueMerge
                                    | _ -> RestartMerge)
                              Shading = tcPr |> Option.bind (fun p -> p.Shading |> opt) |> Option.map colorOfShadingFill
                              Borders =
                                tcPr
                                |> Option.bind (fun p -> p.TableCellBorders |> opt)
                                |> Option.map (fun tcb ->
                                    { Outer =
                                        { Left = tcb.LeftBorder |> opt |> Option.map borderSideOfLeft
                                          Right = tcb.RightBorder |> opt |> Option.map borderSideOfRight
                                          Top = tcb.TopBorder |> opt |> Option.map borderSideOfTop
                                          Bottom = tcb.BottomBorder |> opt |> Option.map borderSideOfBottom }
                                      InsideHorizontal = None
                                      InsideVertical = None })
                              Width =
                                tcPr
                                |> Option.bind (fun p -> p.TableCellWidth |> opt)
                                |> Option.bind (fun w -> w.Width |> opt)
                                |> Option.map (fun v -> twipsToPoints v.Value)
                              Margins =
                                tcPr
                                |> Option.bind (fun p -> p.TableCellMargin |> opt)
                                |> Option.map (fun m ->
                                    { Top = m.TopMargin |> opt |> Option.bind (fun v -> v.Width |> opt) |> Option.map (fun v -> twipsToPoints v.Value)
                                      Bottom = m.BottomMargin |> opt |> Option.bind (fun v -> v.Width |> opt) |> Option.map (fun v -> twipsToPoints v.Value)
                                      Left = m.LeftMargin |> opt |> Option.bind (fun v -> v.Width |> opt) |> Option.map (fun v -> twipsToPoints v.Value)
                                      Right = m.RightMargin |> opt |> Option.bind (fun v -> v.Width |> opt) |> Option.map (fun v -> twipsToPoints v.Value) }) }

                        let content =
                            tc.ChildElements
                            |> Seq.filter (fun c -> not (c :? Wordprocessing.TableCellProperties))
                            |> Seq.collect (readBlock mainPart rctx)
                            |> List.ofSeq

                        { Content = content; Props = props })
                    |> List.ofSeq

                { Cells = cells; Height = height; RepeatAsHeader = repeatAsHeader })
            |> List.ofSeq

        let cellMargins =
            tblPr
            |> Option.bind (fun p -> p.TableCellMarginDefault |> opt)
            |> Option.map (fun m ->
                { Top = m.TopMargin |> opt |> Option.bind (fun v -> v.Width |> opt) |> Option.map (fun v -> twipsToPoints v.Value)
                  Bottom = m.BottomMargin |> opt |> Option.bind (fun v -> v.Width |> opt) |> Option.map (fun v -> twipsToPoints v.Value)
                  Left = m.TableCellLeftMargin |> opt |> Option.bind (fun v -> v.Width |> opt) |> Option.map (fun v -> twipsToPoints (string v.Value))
                  Right = m.TableCellRightMargin |> opt |> Option.bind (fun v -> v.Width |> opt) |> Option.map (fun v -> twipsToPoints (string v.Value)) })

        { Rows = rows
          ColumnWidths = widths
          Style = tableStyleRefOfW tblPr
          Borders = tableBordersOfW tblPr
          CellMargins = cellMargins }

    /// The inverse of `Writer.insertNoteMarker` - strips the note-reference-mark run
    /// (`w:footnoteRef`/`w:endnoteRef`) back out of the body's first paragraph before
    /// handing the rest to `readBlock` like any other content, so a caller reading a
    /// `Footnote`/`Endnote`'s `content` back never sees the marker `Writer` added.
    and private readNoteContent (mainPart: MainDocumentPart) (rctx: ReadCtx) (note: Wordprocessing.FootnoteEndnoteType) : Block list =
        match note.ChildElements |> List.ofSeq with
        | (:? Wordprocessing.Paragraph as firstPara) :: rest ->
            let strippedChildren =
                firstPara.ChildElements
                |> Seq.filter (fun c ->
                    match c with
                    | :? Wordprocessing.Run as r ->
                        (r.Descendants<Wordprocessing.FootnoteReferenceMark>() |> Seq.isEmpty)
                        && (r.Descendants<Wordprocessing.EndnoteReferenceMark>() |> Seq.isEmpty)
                    | _ -> true)
                |> Seq.map (fun c -> c.CloneNode true)

            let strippedPara = Wordprocessing.Paragraph(strippedChildren)
            readBlock mainPart rctx strippedPara @ (rest |> List.collect (readBlock mainPart rctx))
        | other -> other |> List.collect (readBlock mainPart rctx)

    // --- Page setup / headers & footers -------------------------------------------------------

    let private namedPageSizeOfDims (w: int) (h: int) : PageSize =
        match w, h with
        | 12240, 15840 -> Letter
        | 12240, 20160 -> Legal
        | 11906, 16838 -> A4
        | 16838, 23811 -> A3
        | _ -> CustomPageSize(twipsToPoints (string w), twipsToPoints (string h))

    let private pageSizeAndOrientationOfW (ps: Wordprocessing.PageSize option) : PageSize * PageOrientation =
        match ps with
        | None -> Letter, Portrait
        | Some ps ->
            let w = ps.Width |> opt |> Option.map (fun v -> int v.Value) |> Option.defaultValue 12240
            let h = ps.Height |> opt |> Option.map (fun v -> int v.Value) |> Option.defaultValue 15840
            let isLandscape = ps.Orient |> opt |> Option.map (fun v -> v.Value = Wordprocessing.PageOrientationValues.Landscape) |> Option.defaultValue false

            match ps.Code |> opt with
            | Some code -> OtherPageSize(int code.Value), (if isLandscape then Landscape else Portrait)
            | None ->
                let portraitW, portraitH = if isLandscape then h, w else w, h
                namedPageSizeOfDims portraitW portraitH, (if isLandscape then Landscape else Portrait)

    let private pageMarginsOfW (pm: Wordprocessing.PageMargin option) : PageMargins =
        match pm with
        | None -> PageMargins.Default
        | Some pm ->
            { Top = pm.Top |> opt |> Option.map (fun v -> twipsToPoints (string v.Value)) |> Option.defaultValue 72.0
              Bottom = pm.Bottom |> opt |> Option.map (fun v -> twipsToPoints (string v.Value)) |> Option.defaultValue 72.0
              Left = pm.Left |> opt |> Option.map (fun v -> twipsToPoints (string v.Value)) |> Option.defaultValue 72.0
              Right = pm.Right |> opt |> Option.map (fun v -> twipsToPoints (string v.Value)) |> Option.defaultValue 72.0
              Header = pm.Header |> opt |> Option.map (fun v -> twipsToPoints (string v.Value)) |> Option.defaultValue 36.0
              Footer = pm.Footer |> opt |> Option.map (fun v -> twipsToPoints (string v.Value)) |> Option.defaultValue 36.0
              Gutter = pm.Gutter |> opt |> Option.map (fun v -> twipsToPoints (string v.Value)) |> Option.defaultValue 0.0 }

    let private readHeaderPart (mainPart: MainDocumentPart) (rctx: ReadCtx) (id: string) : Block list =
        match mainPart.GetPartById(id) with
        | :? HeaderPart as hp -> hp.Header.ChildElements |> Seq.collect (readBlock mainPart rctx) |> List.ofSeq
        | _ -> []

    let private readFooterPart (mainPart: MainDocumentPart) (rctx: ReadCtx) (id: string) : Block list =
        match mainPart.GetPartById(id) with
        | :? FooterPart as fp -> fp.Footer.ChildElements |> Seq.collect (readBlock mainPart rctx) |> List.ofSeq
        | _ -> []

    let private headerFooterSetOfRefs
        (readPart: string -> Block list)
        (refs: (Wordprocessing.HeaderFooterValues * string) list)
        : HeaderFooterSet option =
        if refs.IsEmpty then
            None
        else
            let find v = refs |> List.tryFind (fun (t, _) -> t = v) |> Option.map (fun (_, id) -> readPart id)

            Some
                { Default = find Wordprocessing.HeaderFooterValues.Default
                  First = find Wordprocessing.HeaderFooterValues.First
                  Even = find Wordprocessing.HeaderFooterValues.Even }

    let private noteNumberRestartOfW (v: Wordprocessing.RestartNumberValues) : NoteNumberRestart =
        if v = Wordprocessing.RestartNumberValues.EachSection then RestartEachSection
        elif v = Wordprocessing.RestartNumberValues.EachPage then RestartEachPage
        else ContinuousRestart

    /// Shared by `FootnoteProperties`/`EndnoteProperties` reading below - both carry the
    /// identically-shaped `NumberingFormat`/`NumberingStart`/`NumberingRestart` triple,
    /// just under different container element types (no common base the SDK exposes this
    /// through, so the caller passes each child in rather than the container itself).
    let private noteNumberingSettingsOfW
        (numFmt: Wordprocessing.NumberingFormat option)
        (numStart: Wordprocessing.NumberingStart option)
        (numRestart: Wordprocessing.NumberingRestart option)
        : NoteNumberingSettings =
        { Format = numFmt |> Option.bind (fun f -> f.Val |> opt) |> Option.map (fun v -> numberFormatKindOfW v.Value) |> Option.defaultValue DecimalFormat
          StartAt = numStart |> Option.bind (fun s -> s.Val |> opt) |> Option.map (fun v -> int v.Value)
          Restart = numRestart |> Option.bind (fun r -> r.Val |> opt) |> Option.map (fun v -> noteNumberRestartOfW v.Value) |> Option.defaultValue ContinuousRestart }

    let private sectionPropertiesOfW (mainPart: MainDocumentPart) (rctx: ReadCtx) (sectPr: Wordprocessing.SectionProperties) : SectionProperties =
        let pageSize, orientation = pageSizeAndOrientationOfW (sectPr.GetFirstChild<Wordprocessing.PageSize>() |> opt)
        let margins = pageMarginsOfW (sectPr.GetFirstChild<Wordprocessing.PageMargin>() |> opt)

        let headerRefs =
            sectPr.Elements<Wordprocessing.HeaderReference>()
            |> Seq.choose (fun r -> match r.Type |> opt, r.Id |> opt with
                                     | Some t, Some id -> Some(t.Value, id.Value)
                                     | _ -> None)
            |> List.ofSeq

        let footerRefs =
            sectPr.Elements<Wordprocessing.FooterReference>()
            |> Seq.choose (fun r -> match r.Type |> opt, r.Id |> opt with
                                     | Some t, Some id -> Some(t.Value, id.Value)
                                     | _ -> None)
            |> List.ofSeq

        let columns =
            sectPr.GetFirstChild<Wordprocessing.Columns>()
            |> opt
            |> Option.bind (fun c -> c.ColumnCount |> opt)
            |> Option.map (fun v -> int v.Value)
            |> Option.defaultValue 1

        let pageNumStart =
            sectPr.GetFirstChild<Wordprocessing.PageNumberType>() |> opt |> Option.bind (fun p -> p.Start |> opt) |> Option.map (fun v -> int v.Value)

        let breakType =
            sectPr.GetFirstChild<Wordprocessing.SectionType>()
            |> opt
            |> Option.bind (fun t -> t.Val |> opt)
            |> Option.map (fun v ->
                if v.Value = Wordprocessing.SectionMarkValues.Continuous then ContinuousBreak
                elif v.Value = Wordprocessing.SectionMarkValues.EvenPage then EvenPageBreak
                elif v.Value = Wordprocessing.SectionMarkValues.OddPage then OddPageBreak
                else NextPageBreak)
            |> Option.defaultValue NextPageBreak

        let footnoteNumbering =
            sectPr.GetFirstChild<Wordprocessing.FootnoteProperties>()
            |> opt
            |> Option.map (fun fpr -> noteNumberingSettingsOfW (fpr.NumberingFormat |> opt) (fpr.NumberingStart |> opt) (fpr.NumberingRestart |> opt))

        let endnoteNumbering =
            sectPr.GetFirstChild<Wordprocessing.EndnoteProperties>()
            |> opt
            |> Option.map (fun epr -> noteNumberingSettingsOfW (epr.NumberingFormat |> opt) (epr.NumberingStart |> opt) (epr.NumberingRestart |> opt))

        { PageSize = pageSize
          Orientation = orientation
          Margins = margins
          Header = headerFooterSetOfRefs (readHeaderPart mainPart rctx) headerRefs
          Footer = headerFooterSetOfRefs (readFooterPart mainPart rctx) footerRefs
          PageNumberStart = pageNumStart
          Columns = columns
          BreakType = breakType
          FootnoteNumbering = footnoteNumbering
          EndnoteNumbering = endnoteNumbering }

    // --- Sections -------------------------------------------------------------------------

    /// The inverse of `Writer.sectionsToBodyChildren` - splits the body's flat child list
    /// back into `Section`s at each embedded `<w:sectPr>` (paragraph-level or, for the last
    /// section, the body's own trailing one).
    let private readSections (mainPart: MainDocumentPart) (rctx: ReadCtx) (body: Wordprocessing.Body) : Section list =
        let children = body.ChildElements |> Array.ofSeq
        let sections = ResizeArray<Section>()
        let currentBlocks = ResizeArray<OpenXmlElement>()

        for child in children do
            match child with
            | :? Wordprocessing.SectionProperties as sectPr ->
                let blocks = currentBlocks |> Seq.collect (readBlock mainPart rctx) |> List.ofSeq
                sections.Add({ Body = blocks; Properties = sectionPropertiesOfW mainPart rctx sectPr })
                currentBlocks.Clear()
            | :? Wordprocessing.Paragraph as p when not (isNull p.ParagraphProperties) && not (isNull p.ParagraphProperties.SectionProperties) ->
                let sectPr = p.ParagraphProperties.SectionProperties
                let strippedPPr = Wordprocessing.ParagraphProperties(p.ParagraphProperties.ChildElements |> Seq.filter (fun c -> not (c :? Wordprocessing.SectionProperties)) |> Seq.map (fun c -> c.CloneNode true))
                let strippedPara = Wordprocessing.Paragraph(p.ChildElements |> Seq.filter (fun c -> not (c :? Wordprocessing.ParagraphProperties)) |> Seq.map (fun c -> c.CloneNode true))

                if strippedPPr.HasChildren then
                    strippedPara.PrependChild(strippedPPr) |> ignore

                currentBlocks.Add(strippedPara)
                let blocks = currentBlocks |> Seq.collect (readBlock mainPart rctx) |> List.ofSeq
                sections.Add({ Body = blocks; Properties = sectionPropertiesOfW mainPart rctx sectPr })
                currentBlocks.Clear()
            | other -> currentBlocks.Add(other)

        if currentBlocks.Count > 0 then
            // No trailing body-level sectPr (malformed/foreign file) - treat remaining
            // content as one final section with default page setup rather than dropping it.
            let blocks = currentBlocks |> Seq.collect (readBlock mainPart rctx) |> List.ofSeq
            sections.Add({ Body = blocks; Properties = SectionProperties.Default })

        sections |> List.ofSeq

    // --- Document protection ---------------------------------------------------------------

    let private editRestrictionOfW (v: Wordprocessing.DocumentProtectionValues) : EditRestriction option =
        if v = Wordprocessing.DocumentProtectionValues.ReadOnly then Some ReadOnlyRestriction
        elif v = Wordprocessing.DocumentProtectionValues.Comments then Some CommentsOnlyRestriction
        elif v = Wordprocessing.DocumentProtectionValues.TrackedChanges then Some TrackedChangesOnlyRestriction
        elif v = Wordprocessing.DocumentProtectionValues.Forms then Some FormsOnlyRestriction
        else None

    let private documentProtectionOfW (settings: Wordprocessing.Settings option) : DocumentProtection option =
        settings
        |> Option.bind (fun s -> s.GetFirstChild<Wordprocessing.DocumentProtection>() |> opt)
        |> Option.bind (fun dp -> dp.Edit |> opt |> Option.bind (fun v -> editRestrictionOfW v.Value))
        |> Option.map (fun edit ->
            // Password never round-trips (the hash isn't reversible) - see
            // `DocumentProtection.Password`'s own doc comment.
            { Edit = Some edit; Password = None })

    // --- Top-level orchestration -------------------------------------------------------------

    let private readDocument (wordDoc: WordprocessingDocument) : Document =
        let mainPart = wordDoc.MainDocumentPart

        let commentsById =
            mainPart.WordprocessingCommentsPart
            |> opt
            |> Option.bind (fun p -> p.Comments |> opt)
            |> Option.map (fun cs -> cs.Elements<Wordprocessing.Comment>() |> Seq.map (fun c -> c.Id.Value, c) |> Map.ofSeq)
            |> Option.defaultValue Map.empty

        // Every `w:bookmarkStart`'s id -> name, from the whole body - a bookmark spanning
        // multiple paragraphs is readable as two independent markers (`BookmarkRangeStart`/
        // `BookmarkRangeEnd`, see `parseInlineRange`), so this can't wait to be built
        // lazily per-paragraph the way a within-paragraph `Bookmark` is.
        let bookmarkNamesById =
            mainPart.Document.Body.Descendants<Wordprocessing.BookmarkStart>()
            |> Seq.map (fun bs -> bs.Id.Value, bs.Name.Value)
            |> Map.ofSeq

        // Note bodies never contain a *nested* footnote/endnote reference (not a real Word
        // feature) - only `CommentsById` needs to be real when reading them, so this
        // bootstraps with empty note maps rather than needing a two-pass/lazy `rctx`.
        let noteReadCtx =
            { CommentsById = commentsById
              FootnotesById = Map.empty
              EndnotesById = Map.empty
              BookmarkNamesById = bookmarkNamesById }

        let readNormalNotes (elements: Wordprocessing.FootnoteEndnoteType seq) =
            elements
            |> Seq.filter (fun n -> n.Type |> opt |> Option.map (fun v -> v.Value = Wordprocessing.FootnoteEndnoteValues.Normal) |> Option.defaultValue true)
            |> Seq.map (fun n -> int n.Id.Value, readNoteContent mainPart noteReadCtx n)
            |> Map.ofSeq

        let footnotesById =
            mainPart.FootnotesPart
            |> opt
            |> Option.bind (fun p -> p.Footnotes |> opt)
            |> Option.map (fun fns -> readNormalNotes (fns.Elements<Wordprocessing.Footnote>() |> Seq.cast<Wordprocessing.FootnoteEndnoteType>))
            |> Option.defaultValue Map.empty

        let endnotesById =
            mainPart.EndnotesPart
            |> opt
            |> Option.bind (fun p -> p.Endnotes |> opt)
            |> Option.map (fun ens -> readNormalNotes (ens.Elements<Wordprocessing.Endnote>() |> Seq.cast<Wordprocessing.FootnoteEndnoteType>))
            |> Option.defaultValue Map.empty

        let rctx =
            { CommentsById = commentsById
              FootnotesById = footnotesById
              EndnotesById = endnotesById
              BookmarkNamesById = bookmarkNamesById }

        let stylesXml = mainPart.StyleDefinitionsPart |> opt |> Option.bind (fun p -> p.Styles |> opt)
        let styles = stylesOfOpenXml stylesXml
        let tableStyles = tableStylesOfOpenXml stylesXml
        let numbering = numberingOfW (mainPart.NumberingDefinitionsPart |> opt |> Option.bind (fun p -> p.Numbering |> opt))
        let protection = documentProtectionOfW (mainPart.DocumentSettingsPart |> opt |> Option.bind (fun p -> p.Settings |> opt))

        let vbaProject =
            mainPart.VbaProjectPart
            |> opt
            |> Option.map (fun p ->
                use stream = p.GetStream()
                use mem = new IO.MemoryStream()
                stream.CopyTo(mem)
                mem.ToArray())

        let nonEmpty (s: string) : string option = s |> Option.ofObj |> Option.filter (fun v -> v <> "")

        let properties =
            { Title = nonEmpty wordDoc.PackageProperties.Title
              Author = nonEmpty wordDoc.PackageProperties.Creator
              Subject = nonEmpty wordDoc.PackageProperties.Subject
              Keywords = nonEmpty wordDoc.PackageProperties.Keywords
              Comments = nonEmpty wordDoc.PackageProperties.Description
              Category = nonEmpty wordDoc.PackageProperties.Category
              Company =
                wordDoc.ExtendedFilePropertiesPart
                |> opt
                |> Option.bind (fun p -> p.Properties |> opt)
                |> Option.bind (fun props -> props.Company |> opt)
                |> Option.bind (fun c -> nonEmpty c.Text) }

        { Sections = readSections mainPart rctx mainPart.Document.Body
          Styles = styles
          Numbering = numbering
          Protection = protection
          VbaProject = vbaProject
          Properties = properties
          TableStyles = tableStyles }

    let loadFromFile (path: string) : Document =
        use wordDoc = WordprocessingDocument.Open(path, false)
        readDocument wordDoc

    let loadFromStream (stream: IO.Stream) : Document =
        use wordDoc = WordprocessingDocument.Open(stream, false)
        readDocument wordDoc
