namespace Kookerella.FsWordDsl

open System
open System.IO
open System.Xml.Linq
open System.Xml.Schema
open System.Reflection

/// A third way in and out of the DSL, alongside writing F# directly and code generation:
/// plain XML, against a real embedded schema (`Xml.xsd`). Mirrors the conventions Excel's
/// own `Xml.fs` documents in its README: a data-carrying DU case becomes an element named
/// after the case (camelCased); a parameterless-choice case becomes an attribute value.
module Xml =

    let private xn (name: string) = XName.Get name
    let private attr (name: string) (v: 'a) = XAttribute(xn name, box v)
    let private attrOpt (name: string) (o: 'a option) : XAttribute list = o |> Option.map (fun v -> attr name v) |> Option.toList

    let private colorToStr (c: Color) : string =
        match c with
        | Rgb(r, g, b) -> sprintf "%02X%02X%02X" r g b
        | Auto -> "auto"

    let private colorOfStr (s: string) : Color =
        if String.Equals(s, "auto", StringComparison.OrdinalIgnoreCase) then
            Auto
        else
            let n = Convert.ToInt32(s, 16)
            Rgb(byte ((n >>> 16) &&& 0xFF), byte ((n >>> 8) &&& 0xFF), byte (n &&& 0xFF))

    let private underlineToStr (u: UnderlineStyle) : string =
        match u with
        | SingleUnderline -> "single"
        | DoubleUnderline -> "double"
        | ThickUnderline -> "thick"
        | DottedUnderline -> "dotted"
        | DashedUnderline -> "dashed"
        | WavyUnderline -> "wavy"
        | OtherUnderline raw -> "other:" + raw

    let private underlineOfStr (s: string) : UnderlineStyle =
        match s with
        | "single" -> SingleUnderline
        | "double" -> DoubleUnderline
        | "thick" -> ThickUnderline
        | "dotted" -> DottedUnderline
        | "dashed" -> DashedUnderline
        | "wavy" -> WavyUnderline
        | other when other.StartsWith("other:") -> OtherUnderline(other.Substring(6))
        | other -> OtherUnderline other

    let private highlightToStr (h: HighlightColor) : string =
        match h with
        | HlYellow -> "yellow"
        | HlGreen -> "green"
        | HlCyan -> "cyan"
        | HlMagenta -> "magenta"
        | HlBlue -> "blue"
        | HlRed -> "red"
        | HlDarkBlue -> "darkBlue"
        | HlDarkCyan -> "darkCyan"
        | HlDarkGreen -> "darkGreen"
        | HlDarkMagenta -> "darkMagenta"
        | HlDarkRed -> "darkRed"
        | HlDarkYellow -> "darkYellow"
        | HlDarkGray -> "darkGray"
        | HlLightGray -> "lightGray"
        | HlBlack -> "black"

    let private highlightOfStr (s: string) : HighlightColor =
        match s with
        | "yellow" -> HlYellow
        | "green" -> HlGreen
        | "cyan" -> HlCyan
        | "magenta" -> HlMagenta
        | "blue" -> HlBlue
        | "red" -> HlRed
        | "darkBlue" -> HlDarkBlue
        | "darkCyan" -> HlDarkCyan
        | "darkGreen" -> HlDarkGreen
        | "darkMagenta" -> HlDarkMagenta
        | "darkRed" -> HlDarkRed
        | "darkYellow" -> HlDarkYellow
        | "darkGray" -> HlDarkGray
        | "lightGray" -> HlLightGray
        | _ -> HlBlack

    let private vertPosToStr (v: VerticalPosition) : string =
        match v with
        | Superscript -> "superscript"
        | Subscript -> "subscript"

    let private vertPosOfStr (s: string) : VerticalPosition = if s = "subscript" then Subscript else Superscript

    let private alignToStr (a: ParagraphAlignment) : string =
        match a with
        | AlignLeft -> "left"
        | AlignCenter -> "center"
        | AlignRight -> "right"
        | AlignJustify -> "justify"

    let private alignOfStr (s: string) : ParagraphAlignment =
        match s with
        | "center" -> AlignCenter
        | "right" -> AlignRight
        | "justify" -> AlignJustify
        | _ -> AlignLeft

    let private lineSpacingToXml (ls: LineSpacingRule) : XElement =
        match ls with
        | SingleSpacing -> XElement(xn "lineSpacing", attr "kind" "single")
        | OnePointFiveSpacing -> XElement(xn "lineSpacing", attr "kind" "onePointFive")
        | DoubleSpacing -> XElement(xn "lineSpacing", attr "kind" "double")
        | AtLeastSpacing p -> XElement(xn "lineSpacing", attr "kind" "atLeast", attr "points" p)
        | ExactlySpacing p -> XElement(xn "lineSpacing", attr "kind" "exactly", attr "points" p)
        | MultipleSpacing f -> XElement(xn "lineSpacing", attr "kind" "multiple", attr "factor" f)

    let private lineSpacingOfXml (el: XElement) : LineSpacingRule =
        let kind = el.Attribute(xn "kind").Value
        let pts () = float (el.Attribute(xn "points").Value)
        let factor () = float (el.Attribute(xn "factor").Value)

        match kind with
        | "onePointFive" -> OnePointFiveSpacing
        | "double" -> DoubleSpacing
        | "atLeast" -> AtLeastSpacing(pts ())
        | "exactly" -> ExactlySpacing(pts ())
        | "multiple" -> MultipleSpacing(factor ())
        | _ -> SingleSpacing

    let private borderLineToStr (s: BorderLineStyle) : string =
        match s with
        | SingleLine -> "single"
        | ThickLine -> "thick"
        | DoubleLine -> "double"
        | DottedLine -> "dotted"
        | DashedLine -> "dashed"
        | WaveLine -> "wave"
        | OtherLine raw -> "other:" + raw

    let private borderLineOfStr (s: string) : BorderLineStyle =
        match s with
        | "single" -> SingleLine
        | "thick" -> ThickLine
        | "double" -> DoubleLine
        | "dotted" -> DottedLine
        | "dashed" -> DashedLine
        | "wave" -> WaveLine
        | other when other.StartsWith("other:") -> OtherLine(other.Substring(6))
        | other -> OtherLine other

    // --- Run / paragraph formatting ---------------------------------------------------------

    let private runStyleToXml (s: RunStyle) : XElement =
        XElement(
            xn "runStyle",
            attrOpt "fontFamily" s.FontFamily,
            attrOpt "size" s.Size,
            (if s.Bold then [ attr "bold" true ] else []),
            (if s.Italic then [ attr "italic" true ] else []),
            attrOpt "underline" (s.Underline |> Option.map underlineToStr),
            (if s.Strikethrough then [ attr "strikethrough" true ] else []),
            attrOpt "color" (s.Color |> Option.map colorToStr),
            attrOpt "highlight" (s.Highlight |> Option.map highlightToStr),
            attrOpt "verticalPosition" (s.VerticalPosition |> Option.map vertPosToStr)
        )

    let private strAttr (name: string) (el: XElement) : string option = el.Attribute(xn name) |> Option.ofObj |> Option.map (fun a -> a.Value)
    let private boolAttr (name: string) (el: XElement) : bool = strAttr name el |> Option.map bool.Parse |> Option.defaultValue false

    let private runStyleOfXml (el: XElement) : RunStyle =
        { FontFamily = strAttr "fontFamily" el
          Size = strAttr "size" el |> Option.map float
          Bold = boolAttr "bold" el
          Italic = boolAttr "italic" el
          Underline = strAttr "underline" el |> Option.map underlineOfStr
          Strikethrough = boolAttr "strikethrough" el
          Color = strAttr "color" el |> Option.map colorOfStr
          Highlight = strAttr "highlight" el |> Option.map highlightOfStr
          VerticalPosition = strAttr "verticalPosition" el |> Option.map vertPosOfStr }

    let private indentationToXml (i: Indentation) : XElement =
        XElement(xn "indentation", attrOpt "left" i.Left, attrOpt "right" i.Right, attrOpt "firstLine" i.FirstLine, attrOpt "hanging" i.Hanging)

    let private indentationOfXml (el: XElement) : Indentation =
        { Left = strAttr "left" el |> Option.map float
          Right = strAttr "right" el |> Option.map float
          FirstLine = strAttr "firstLine" el |> Option.map float
          Hanging = strAttr "hanging" el |> Option.map float }

    let private paragraphFormatToXml (f: ParagraphFormat) : XElement =
        XElement(
            xn "paragraphFormat",
            attrOpt "alignment" (f.Alignment |> Option.map alignToStr),
            attrOpt "spacingBefore" f.SpacingBefore,
            attrOpt "spacingAfter" f.SpacingAfter,
            (if f.KeepWithNext then [ attr "keepWithNext" true ] else []),
            (if f.PageBreakBefore then [ attr "pageBreakBefore" true ] else []),
            (f.LineSpacing |> Option.map lineSpacingToXml |> Option.toList),
            (f.Indentation |> Option.map indentationToXml |> Option.toList)
        )

    let private paragraphFormatOfXml (el: XElement) : ParagraphFormat =
        { Alignment = strAttr "alignment" el |> Option.map alignOfStr
          SpacingBefore = strAttr "spacingBefore" el |> Option.map float
          SpacingAfter = strAttr "spacingAfter" el |> Option.map float
          LineSpacing = el.Element(xn "lineSpacing") |> Option.ofObj |> Option.map lineSpacingOfXml
          Indentation = el.Element(xn "indentation") |> Option.ofObj |> Option.map indentationOfXml
          KeepWithNext = boolAttr "keepWithNext" el
          PageBreakBefore = boolAttr "pageBreakBefore" el }

    // --- Borders --------------------------------------------------------------------------

    let private borderSideToXml (name: string) (s: BorderSide) : XElement =
        XElement(xn name, attr "style" (borderLineToStr s.Style), attrOpt "width" s.Width, attrOpt "color" (s.Color |> Option.map colorToStr))

    let private borderSideOfXml (el: XElement) : BorderSide =
        { Style = borderLineOfStr (el.Attribute(xn "style").Value)
          Width = strAttr "width" el |> Option.map float
          Color = strAttr "color" el |> Option.map colorOfStr }

    let private borderStyleChildren (b: BorderStyle) : XElement list =
        [ b.Left |> Option.map (borderSideToXml "left")
          b.Right |> Option.map (borderSideToXml "right")
          b.Top |> Option.map (borderSideToXml "top")
          b.Bottom |> Option.map (borderSideToXml "bottom") ]
        |> List.choose id

    let private borderStyleOfXml (el: XElement) : BorderStyle =
        { Left = el.Element(xn "left") |> Option.ofObj |> Option.map borderSideOfXml
          Right = el.Element(xn "right") |> Option.ofObj |> Option.map borderSideOfXml
          Top = el.Element(xn "top") |> Option.ofObj |> Option.map borderSideOfXml
          Bottom = el.Element(xn "bottom") |> Option.ofObj |> Option.map borderSideOfXml }

    let private tableBordersToXml (b: TableBorders) : XElement =
        XElement(
            xn "tableBorders",
            borderStyleChildren b.Outer,
            (b.InsideHorizontal |> Option.map (borderSideToXml "insideHorizontal") |> Option.toList),
            (b.InsideVertical |> Option.map (borderSideToXml "insideVertical") |> Option.toList)
        )

    let private tableBordersOfXml (el: XElement) : TableBorders =
        { Outer = borderStyleOfXml el
          InsideHorizontal = el.Element(xn "insideHorizontal") |> Option.ofObj |> Option.map borderSideOfXml
          InsideVertical = el.Element(xn "insideVertical") |> Option.ofObj |> Option.map borderSideOfXml }

    // --- Images / hyperlinks -----------------------------------------------------------------

    let private imageToXml (img: ImageEntry) : XElement =
        XElement(
            xn "image",
            attr "format" (sprintf "%A" img.Format),
            attr "widthEmu" img.WidthEmu,
            attr "heightEmu" img.HeightEmu,
            attrOpt "altText" img.AltText,
            Convert.ToBase64String(img.Data)
        )

    let private imageOfXml (el: XElement) : ImageEntry =
        let format =
            match el.Attribute(xn "format").Value with
            | "Jpeg" -> Jpeg
            | "Gif" -> Gif
            | "Bmp" -> Bmp
            | _ -> Png

        { Data = Convert.FromBase64String(el.Value)
          Format = format
          WidthEmu = int64 (el.Attribute(xn "widthEmu").Value)
          HeightEmu = int64 (el.Attribute(xn "heightEmu").Value)
          AltText = strAttr "altText" el }

    let private hyperlinkTargetToXml (t: HyperlinkTarget) : XElement =
        match t with
        | ExternalUrl u -> XElement(xn "externalHyperlink", u)
        | InternalBookmark n -> XElement(xn "internalHyperlink", n)

    let private hyperlinkTargetOfXml (el: XElement) : HyperlinkTarget =
        if el.Name.LocalName = "internalHyperlink" then InternalBookmark el.Value else ExternalUrl el.Value

    // --- Inline content ---------------------------------------------------------------------

    let private tableStyleRefToXml (s: TableStyleRef) : XElement =
        XElement(
            xn "tableStyle",
            attr "name" s.Name,
            (if s.FirstRowBanding then [ attr "firstRow" true ] else []),
            (if s.LastRowBanding then [ attr "lastRow" true ] else []),
            (if s.BandedRows then [ attr "bandedRows" true ] else []),
            (if s.BandedColumns then [ attr "bandedColumns" true ] else [])
        )

    let private tableStyleRefOfXml (el: XElement) : TableStyleRef =
        { Name = el.Attribute(xn "name").Value
          FirstRowBanding = boolAttr "firstRow" el
          LastRowBanding = boolAttr "lastRow" el
          BandedRows = boolAttr "bandedRows" el
          BandedColumns = boolAttr "bandedColumns" el }

    let private tableCellPropsToXml (p: TableCellProps) : XElement =
        XElement(
            xn "cellProps",
            attrOpt "gridSpan" p.GridSpan,
            attrOpt "verticalMerge" (p.VerticalMerge |> Option.map (function RestartMerge -> "restart" | ContinueMerge -> "continue")),
            attrOpt "shading" (p.Shading |> Option.map colorToStr),
            attrOpt "width" p.Width,
            (p.Borders |> Option.map tableBordersToXml |> Option.toList)
        )

    let private tableCellPropsOfXml (el: XElement) : TableCellProps =
        { GridSpan = strAttr "gridSpan" el |> Option.map int
          VerticalMerge = strAttr "verticalMerge" el |> Option.map (fun s -> if s = "continue" then ContinueMerge else RestartMerge)
          Shading = strAttr "shading" el |> Option.map colorOfStr
          Borders = el.Element(xn "tableBorders") |> Option.ofObj |> Option.map tableBordersOfXml
          Width = strAttr "width" el |> Option.map float }

    // `inlineToXml`/`inlineOfXml` need `blockToXml`/`blockOfXml` (a `Footnote`/`Endnote`'s
    // own body is a `Block list`), which need `paragraphToXml`/`paragraphOfXml`, which need
    // `inlineToXml`/`inlineOfXml` back for a paragraph's own `Inlines` - one `rec ... and
    // ...` chain for the same reason `Writer.fs`'s equivalent functions are chained.
    let rec private inlineToXml (i: Inline) : XElement =
        match i with
        | Run(text, style, styleId) ->
            XElement(xn "run", attrOpt "styleId" styleId, (style |> Option.map runStyleToXml |> Option.toList), text)
        | LineBreak -> XElement(xn "lineBreak")
        | Tab -> XElement(xn "tab")
        | PageBreak -> XElement(xn "pageBreak")
        | Image img -> imageToXml img
        | Hyperlink(target, runs, tooltip) ->
            XElement(xn "hyperlink", attrOpt "tooltip" tooltip, hyperlinkTargetToXml target, XElement(xn "content", runs |> List.map inlineToXml))
        | Bookmark(name, content) -> XElement(xn "bookmark", attr "name" name, content |> List.map inlineToXml)
        | Comment(author, initials, date, text, content) ->
            XElement(
                xn "comment",
                attr "author" author,
                attrOpt "initials" initials,
                attrOpt "date" (date |> Option.map (fun d -> d.ToString("o"))),
                XElement(xn "text", text),
                XElement(xn "content", content |> List.map inlineToXml)
            )
        | Field(instr, cached) -> XElement(xn "field", attr "instruction" instr, attrOpt "cachedResult" cached)
        | Footnote content -> XElement(xn "footnote", XElement(xn "body", content |> List.map blockToXml))
        | Endnote content -> XElement(xn "endnote", XElement(xn "body", content |> List.map blockToXml))

    and private inlineOfXml (el: XElement) : Inline =
        match el.Name.LocalName with
        | "run" -> Run(el.Value, el.Element(xn "runStyle") |> Option.ofObj |> Option.map runStyleOfXml, strAttr "styleId" el)
        | "lineBreak" -> LineBreak
        | "tab" -> Tab
        | "pageBreak" -> PageBreak
        | "image" -> Image(imageOfXml el)
        | "hyperlink" ->
            let target = el.Elements() |> Seq.find (fun c -> c.Name.LocalName = "externalHyperlink" || c.Name.LocalName = "internalHyperlink") |> hyperlinkTargetOfXml
            let content = el.Element(xn "content").Elements() |> Seq.map inlineOfXml |> List.ofSeq
            Hyperlink(target, content, strAttr "tooltip" el)
        | "bookmark" -> Bookmark(el.Attribute(xn "name").Value, el.Elements() |> Seq.map inlineOfXml |> List.ofSeq)
        | "comment" ->
            let content = el.Element(xn "content").Elements() |> Seq.map inlineOfXml |> List.ofSeq
            let text = el.Element(xn "text").Value
            Comment(el.Attribute(xn "author").Value, strAttr "initials" el, strAttr "date" el |> Option.map DateTime.Parse, text, content)
        | "field" -> Field(el.Attribute(xn "instruction").Value, strAttr "cachedResult" el)
        | "footnote" -> Footnote(el.Element(xn "body").Elements() |> Seq.map blockOfXml |> List.ofSeq)
        | "endnote" -> Endnote(el.Element(xn "body").Elements() |> Seq.map blockOfXml |> List.ofSeq)
        | other -> failwithf "Unknown inline element: %s" other

    and private paragraphToXml (p: Paragraph) : XElement =
        XElement(
            xn "para",
            attrOpt "styleId" p.StyleId,
            attrOpt "numId" (p.Numbering |> Option.map fst),
            attrOpt "level" (p.Numbering |> Option.map snd),
            (p.Format |> Option.map paragraphFormatToXml |> Option.toList),
            p.Inlines |> List.map inlineToXml
        )

    and private paragraphOfXml (el: XElement) : Paragraph =
        let numId = strAttr "numId" el |> Option.map int
        let level = strAttr "level" el |> Option.map int

        { Inlines = el.Elements() |> Seq.filter (fun c -> c.Name.LocalName <> "paragraphFormat") |> Seq.map inlineOfXml |> List.ofSeq
          StyleId = strAttr "styleId" el
          Format = el.Element(xn "paragraphFormat") |> Option.ofObj |> Option.map paragraphFormatOfXml
          Numbering = match numId, level with
                      | Some n, Some l -> Some(n, l)
                      | _ -> None }

    and private blockToXml (b: Block) : XElement =
        match b with
        | ParagraphBlock p -> paragraphToXml p
        | TableBlock t -> tableToXml t

    and private blockOfXml (el: XElement) : Block =
        match el.Name.LocalName with
        | "para" -> ParagraphBlock(paragraphOfXml el)
        | "table" -> TableBlock(tableOfXml el)
        | other -> failwithf "Unknown block element: %s" other

    and private tableCellToXml (c: TableCell) : XElement =
        XElement(xn "cell", tableCellPropsToXml c.Props, c.Content |> List.map blockToXml)

    and private tableCellOfXml (el: XElement) : TableCell =
        { Content = el.Elements() |> Seq.filter (fun c -> c.Name.LocalName <> "cellProps") |> Seq.map blockOfXml |> List.ofSeq
          Props = el.Element(xn "cellProps") |> Option.ofObj |> Option.map tableCellPropsOfXml |> Option.defaultValue TableCellProps.Default }

    and private tableRowToXml (r: TableRow) : XElement = XElement(xn "row", attrOpt "height" r.Height, r.Cells |> List.map tableCellToXml)

    and private tableRowOfXml (el: XElement) : TableRow =
        { Cells = el.Elements(xn "cell") |> Seq.map tableCellOfXml |> List.ofSeq
          Height = strAttr "height" el |> Option.map float }

    and private tableToXml (t: TableEntry) : XElement =
        XElement(
            xn "table",
            (t.Style |> Option.map tableStyleRefToXml |> Option.toList),
            (t.Borders |> Option.map tableBordersToXml |> Option.toList),
            XElement(xn "columnWidths", t.ColumnWidths |> List.map (fun w -> XElement(xn "col", attr "width" w))),
            XElement(xn "rows", t.Rows |> List.map tableRowToXml)
        )

    and private tableOfXml (el: XElement) : TableEntry =
        { Rows = el.Element(xn "rows").Elements(xn "row") |> Seq.map tableRowOfXml |> List.ofSeq
          ColumnWidths = el.Element(xn "columnWidths").Elements(xn "col") |> Seq.map (fun c -> float (c.Attribute(xn "width").Value)) |> List.ofSeq
          Style = el.Element(xn "tableStyle") |> Option.ofObj |> Option.map tableStyleRefOfXml
          Borders = el.Element(xn "tableBorders") |> Option.ofObj |> Option.map tableBordersOfXml }

    // --- Page setup / headers & footers -------------------------------------------------------

    let private pageSizeToXml (p: PageSize) : XElement =
        match p with
        | OtherPageSize code -> XElement(xn "pageSize", attr "other" code)
        | CustomPageSize(w, h) -> XElement(xn "pageSize", attr "widthPoints" w, attr "heightPoints" h)
        | named -> XElement(xn "pageSize", attr "kind" (sprintf "%A" named))

    let private pageSizeOfXml (el: XElement) : PageSize =
        match strAttr "other" el, strAttr "widthPoints" el, strAttr "heightPoints" el, strAttr "kind" el with
        | Some code, _, _, _ -> OtherPageSize(int code)
        | _, Some w, Some h, _ -> CustomPageSize(float w, float h)
        | _, _, _, Some "Legal" -> Legal
        | _, _, _, Some "A4" -> A4
        | _, _, _, Some "A3" -> A3
        | _ -> Letter

    let private pageMarginsToXml (m: PageMargins) : XElement =
        XElement(xn "margins", attr "top" m.Top, attr "bottom" m.Bottom, attr "left" m.Left, attr "right" m.Right, attr "header" m.Header, attr "footer" m.Footer, attr "gutter" m.Gutter)

    let private pageMarginsOfXml (el: XElement) : PageMargins =
        { Top = float (el.Attribute(xn "top").Value)
          Bottom = float (el.Attribute(xn "bottom").Value)
          Left = float (el.Attribute(xn "left").Value)
          Right = float (el.Attribute(xn "right").Value)
          Header = float (el.Attribute(xn "header").Value)
          Footer = float (el.Attribute(xn "footer").Value)
          Gutter = float (el.Attribute(xn "gutter").Value) }

    let rec private headerFooterSetToXml (name: string) (h: HeaderFooterSet) : XElement =
        let variant (n: string) (blocks: Block list option) =
            blocks |> Option.map (fun bs -> XElement(xn n, bs |> List.map blockToXml)) |> Option.toList

        XElement(xn name, variant "default" h.Default, variant "first" h.First, variant "even" h.Even)

    let private headerFooterSetOfXml (el: XElement) : HeaderFooterSet =
        let variant n = el.Element(xn n) |> Option.ofObj |> Option.map (fun e -> e.Elements() |> Seq.map blockOfXml |> List.ofSeq)
        { Default = variant "default"; First = variant "first"; Even = variant "even" }

    let private sectionPropertiesToXml (s: SectionProperties) : XElement =
        XElement(
            xn "pageSetup",
            attr "orientation" (sprintf "%A" s.Orientation),
            attrOpt "pageNumberStart" s.PageNumberStart,
            attr "columns" s.Columns,
            attr "breakType" (sprintf "%A" s.BreakType),
            pageSizeToXml s.PageSize,
            pageMarginsToXml s.Margins,
            (s.Header |> Option.map (headerFooterSetToXml "header") |> Option.toList),
            (s.Footer |> Option.map (headerFooterSetToXml "footer") |> Option.toList)
        )

    let private sectionBreakTypeOfXml (s: string option) : SectionBreakType =
        match s with
        | Some "ContinuousBreak" -> ContinuousBreak
        | Some "EvenPageBreak" -> EvenPageBreak
        | Some "OddPageBreak" -> OddPageBreak
        | _ -> NextPageBreak

    let private sectionPropertiesOfXml (el: XElement) : SectionProperties =
        { PageSize = pageSizeOfXml (el.Element(xn "pageSize"))
          Orientation = (if el.Attribute(xn "orientation").Value = "Landscape" then Landscape else Portrait)
          Margins = pageMarginsOfXml (el.Element(xn "margins"))
          Header = el.Element(xn "header") |> Option.ofObj |> Option.map headerFooterSetOfXml
          Footer = el.Element(xn "footer") |> Option.ofObj |> Option.map headerFooterSetOfXml
          PageNumberStart = strAttr "pageNumberStart" el |> Option.map int
          Columns = int (el.Attribute(xn "columns").Value)
          BreakType = sectionBreakTypeOfXml (strAttr "breakType" el) }

    let private sectionToXml (s: Section) : XElement =
        XElement(xn "section", sectionPropertiesToXml s.Properties, XElement(xn "body", s.Body |> List.map blockToXml))

    let private sectionOfXml (el: XElement) : Section =
        { Body = el.Element(xn "body").Elements() |> Seq.map blockOfXml |> List.ofSeq
          Properties = sectionPropertiesOfXml (el.Element(xn "pageSetup")) }

    // --- Styles / numbering / protection -----------------------------------------------------

    let private styleDefinitionToXml (d: StyleDefinition) : XElement =
        XElement(
            xn "style",
            attr "id" d.Id,
            attr "name" d.Name,
            attr "type" (match d.Type with ParagraphStyleType -> "paragraph" | CharacterStyleType -> "character"),
            attrOpt "basedOn" d.BasedOn,
            (d.RunFormat |> Option.map runStyleToXml |> Option.toList),
            (d.ParaFormat |> Option.map paragraphFormatToXml |> Option.toList)
        )

    let private styleDefinitionOfXml (el: XElement) : StyleDefinition =
        { Id = el.Attribute(xn "id").Value
          Name = el.Attribute(xn "name").Value
          Type = if el.Attribute(xn "type").Value = "character" then CharacterStyleType else ParagraphStyleType
          BasedOn = strAttr "basedOn" el
          RunFormat = el.Element(xn "runStyle") |> Option.ofObj |> Option.map runStyleOfXml
          ParaFormat = el.Element(xn "paragraphFormat") |> Option.ofObj |> Option.map paragraphFormatOfXml }

    let private numberFormatKindToXml (k: NumberFormatKind) : XElement =
        match k with
        | BulletFormat(glyph, font) -> XElement(xn "format", attr "kind" "bullet", attr "glyph" (int glyph), attr "font" font)
        | OtherFormat raw -> XElement(xn "format", attr "kind" "other", attr "raw" raw)
        | other -> XElement(xn "format", attr "kind" (sprintf "%A" other))

    let private numberFormatKindOfXml (el: XElement) : NumberFormatKind =
        match el.Attribute(xn "kind").Value with
        | "bullet" -> BulletFormat(char (int (el.Attribute(xn "glyph").Value)), el.Attribute(xn "font").Value)
        | "other" -> OtherFormat(el.Attribute(xn "raw").Value)
        | "LowerLetterFormat" -> LowerLetterFormat
        | "UpperLetterFormat" -> UpperLetterFormat
        | "LowerRomanFormat" -> LowerRomanFormat
        | "UpperRomanFormat" -> UpperRomanFormat
        | _ -> DecimalFormat

    let private listLevelToXml (l: ListLevel) : XElement =
        XElement(xn "level", numberFormatKindToXml l.Format, attr "text" l.Text, attrOpt "indentLeft" l.IndentLeft, attrOpt "hangingIndent" l.HangingIndent, attrOpt "startAt" l.StartAt)

    let private listLevelOfXml (el: XElement) : ListLevel =
        { Format = numberFormatKindOfXml (el.Element(xn "format"))
          Text = el.Attribute(xn "text").Value
          IndentLeft = strAttr "indentLeft" el |> Option.map float
          HangingIndent = strAttr "hangingIndent" el |> Option.map float
          StartAt = strAttr "startAt" el |> Option.map int }

    let private numberingDefinitionToXml (d: NumberingDefinition) : XElement =
        XElement(xn "numberingDef", attr "id" d.Id, d.Levels |> List.map listLevelToXml)

    let private numberingDefinitionOfXml (el: XElement) : NumberingDefinition =
        { Id = int (el.Attribute(xn "id").Value)
          Levels = el.Elements(xn "level") |> Seq.map listLevelOfXml |> List.ofSeq }

    let private protectionToXml (p: DocumentProtection) : XElement =
        XElement(xn "protection", attrOpt "edit" (p.Edit |> Option.map (sprintf "%A")), attrOpt "password" p.Password)

    let private protectionOfXml (el: XElement) : DocumentProtection =
        let edit =
            strAttr "edit" el
            |> Option.map (function
                | "CommentsOnlyRestriction" -> CommentsOnlyRestriction
                | "TrackedChangesOnlyRestriction" -> TrackedChangesOnlyRestriction
                | "FormsOnlyRestriction" -> FormsOnlyRestriction
                | _ -> ReadOnlyRestriction)

        { Edit = edit; Password = strAttr "password" el }

    // --- Top level ------------------------------------------------------------------------

    /// `Document` -> `XElement`. See this file's own conventions above and the worked
    /// examples in the root README.
    let toDocument (doc: Document) : XElement =
        XElement(
            xn "document",
            XElement(xn "sections", doc.Sections |> List.map sectionToXml),
            (if doc.Styles.IsEmpty then [] else [ XElement(xn "styles", doc.Styles |> List.map styleDefinitionToXml) ]),
            (if doc.Numbering.IsEmpty then [] else [ XElement(xn "numbering", doc.Numbering |> List.map numberingDefinitionToXml) ]),
            (doc.Protection |> Option.map protectionToXml |> Option.toList),
            (doc.VbaProject |> Option.map (fun b -> XElement(xn "vbaProject", Convert.ToBase64String(b))) |> Option.toList)
        )

    /// `XElement` -> `Document`, the inverse of `toDocument`.
    let ofDocument (el: XElement) : Document =
        { Sections = el.Element(xn "sections").Elements(xn "section") |> Seq.map sectionOfXml |> List.ofSeq
          Styles = el.Element(xn "styles") |> Option.ofObj |> Option.map (fun e -> e.Elements(xn "style") |> Seq.map styleDefinitionOfXml |> List.ofSeq) |> Option.defaultValue []
          Numbering = el.Element(xn "numbering") |> Option.ofObj |> Option.map (fun e -> e.Elements(xn "numberingDef") |> Seq.map numberingDefinitionOfXml |> List.ofSeq) |> Option.defaultValue []
          Protection = el.Element(xn "protection") |> Option.ofObj |> Option.map protectionOfXml
          VbaProject = el.Element(xn "vbaProject") |> Option.ofObj |> Option.map (fun e -> Convert.FromBase64String(e.Value)) }

    /// Loads the embedded schema for validating either direction yourself
    /// (`XDocument.Validate`).
    let schemaSet () : XmlSchemaSet =
        let assembly = Assembly.GetExecutingAssembly()
        use stream = assembly.GetManifestResourceStream("Kookerella.FsWordDsl.Xml.xsd")
        let schema = XmlSchema.Read(stream, null)
        let set = XmlSchemaSet()
        set.Add(schema) |> ignore
        set.Compile()
        set
