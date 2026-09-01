namespace Kookerella.FsWordDsl

open System
open System.Text.Json.Nodes

/// A fourth way in and out of the DSL, alongside writing F# directly, code generation, and
/// XML: plain JSON, for a caller whose tooling speaks JSON rather than XML. Same
/// worksheet/workbook-level... er, section/document-level feature set `Xml.fs` covers, same
/// DU-case conventions (a data-carrying case becomes a single-key object named after the
/// case; a parameterless-choice case becomes a bare string). Schema validation
/// (`Json.schema.json`) is test-suite only, same posture as Excel's own `Json.fs`.
module Json =

    let private obj_ (pairs: (string * JsonNode option) list) : JsonObject =
        let o = JsonObject()
        for (k, v) in pairs do
            v |> Option.iter (fun v -> o.[k] <- v)
        o

    let private jstr (s: string) : JsonNode = JsonValue.Create(s)
    let private jnum (n: float) : JsonNode = JsonValue.Create(n)
    let private jint (n: int) : JsonNode = JsonValue.Create(n)
    let private jbool (b: bool) : JsonNode = JsonValue.Create(b)

    let private prop (name: string) (o: JsonObject) : JsonObject option =
        match o.[name] with
        | null -> None
        | n -> Some(n.AsObject())

    let private str (name: string) (o: JsonObject) : string option =
        match o.[name] with
        | null -> None
        | n -> Some(n.GetValue<string>())

    let private num (name: string) (o: JsonObject) : float option =
        match o.[name] with
        | null -> None
        | n -> Some(n.GetValue<float>())

    let private intg (name: string) (o: JsonObject) : int option =
        match o.[name] with
        | null -> None
        | n -> Some(n.GetValue<int>())

    let private boolean (name: string) (o: JsonObject) : bool =
        match o.[name] with
        | null -> false
        | n -> n.GetValue<bool>()

    let private arr (name: string) (o: JsonObject) : JsonNode list =
        match o.[name] with
        | null -> []
        | n -> n.AsArray() |> List.ofSeq

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

    let private highlightToStr (h: HighlightColor) : string = (sprintf "%A" h).Substring(2) |> fun s -> Char.ToLowerInvariant(s.[0]).ToString() + s.Substring(1)

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

    let private runStyleToJson (s: RunStyle) : JsonNode =
        obj_
            [ "fontFamily", s.FontFamily |> Option.map jstr
              "size", s.Size |> Option.map jnum
              "bold", (if s.Bold then Some(jbool true) else None)
              "italic", (if s.Italic then Some(jbool true) else None)
              "underline", s.Underline |> Option.map (underlineToStr >> jstr)
              "strikethrough", (if s.Strikethrough then Some(jbool true) else None)
              "color", s.Color |> Option.map (colorToStr >> jstr)
              "highlight", s.Highlight |> Option.map (highlightToStr >> jstr)
              "verticalPosition", s.VerticalPosition |> Option.map (fun v -> jstr (match v with Superscript -> "superscript" | Subscript -> "subscript"))
              "smallCaps", (if s.SmallCaps then Some(jbool true) else None)
              "allCaps", (if s.AllCaps then Some(jbool true) else None)
              "hidden", (if s.Hidden then Some(jbool true) else None) ]
        :> JsonNode

    let private runStyleOfJson (o: JsonObject) : RunStyle =
        { FontFamily = str "fontFamily" o
          Size = num "size" o
          Bold = boolean "bold" o
          Italic = boolean "italic" o
          Underline = str "underline" o |> Option.map underlineOfStr
          Strikethrough = boolean "strikethrough" o
          Color = str "color" o |> Option.map colorOfStr
          Highlight = str "highlight" o |> Option.map highlightOfStr
          VerticalPosition = str "verticalPosition" o |> Option.map (fun s -> if s = "subscript" then Subscript else Superscript)
          SmallCaps = boolean "smallCaps" o
          AllCaps = boolean "allCaps" o
          Hidden = boolean "hidden" o }

    let private indentationToJson (i: Indentation) : JsonNode =
        obj_ [ "left", i.Left |> Option.map jnum; "right", i.Right |> Option.map jnum; "firstLine", i.FirstLine |> Option.map jnum; "hanging", i.Hanging |> Option.map jnum ]
        :> JsonNode

    let private indentationOfJson (o: JsonObject) : Indentation =
        { Left = num "left" o; Right = num "right" o; FirstLine = num "firstLine" o; Hanging = num "hanging" o }

    let private lineSpacingToJson (ls: LineSpacingRule) : JsonNode =
        match ls with
        | SingleSpacing -> jstr "single"
        | OnePointFiveSpacing -> jstr "onePointFive"
        | DoubleSpacing -> jstr "double"
        | AtLeastSpacing p -> obj_ [ "atLeast", Some(jnum p) ] :> JsonNode
        | ExactlySpacing p -> obj_ [ "exactly", Some(jnum p) ] :> JsonNode
        | MultipleSpacing f -> obj_ [ "multiple", Some(jnum f) ] :> JsonNode

    let private lineSpacingOfJson (n: JsonNode) : LineSpacingRule =
        match n with
        | :? JsonValue as v ->
            match v.GetValue<string>() with
            | "onePointFive" -> OnePointFiveSpacing
            | "double" -> DoubleSpacing
            | _ -> SingleSpacing
        | _ ->
            let o = n.AsObject()
            if not (isNull o.["atLeast"]) then AtLeastSpacing(o.["atLeast"].GetValue<float>())
            elif not (isNull o.["exactly"]) then ExactlySpacing(o.["exactly"].GetValue<float>())
            else MultipleSpacing(o.["multiple"].GetValue<float>())

    // --- Borders --------------------------------------------------------------------------

    let private borderSideToJson (s: BorderSide) : JsonNode =
        obj_ [ "style", Some(jstr (borderLineToStr s.Style)); "width", s.Width |> Option.map jnum; "color", s.Color |> Option.map (colorToStr >> jstr) ] :> JsonNode

    let private borderSideOfJson (o: JsonObject) : BorderSide =
        { Style = borderLineOfStr (str "style" o |> Option.get); Width = num "width" o; Color = str "color" o |> Option.map colorOfStr }

    let private borderStyleToJson (b: BorderStyle) : JsonNode =
        obj_
            [ "left", b.Left |> Option.map borderSideToJson
              "right", b.Right |> Option.map borderSideToJson
              "top", b.Top |> Option.map borderSideToJson
              "bottom", b.Bottom |> Option.map borderSideToJson ]
        :> JsonNode

    let private borderStyleOfJson (o: JsonObject) : BorderStyle =
        { Left = prop "left" o |> Option.map borderSideOfJson
          Right = prop "right" o |> Option.map borderSideOfJson
          Top = prop "top" o |> Option.map borderSideOfJson
          Bottom = prop "bottom" o |> Option.map borderSideOfJson }

    let private tabAlignToStr (a: TabStopAlignment) : string =
        match a with
        | LeftTab -> "left"
        | CenterTab -> "center"
        | RightTab -> "right"
        | DecimalTab -> "decimal"
        | BarTab -> "bar"
        | OtherTabAlignment raw -> "other:" + raw

    let private tabAlignOfStr (s: string) : TabStopAlignment =
        match s with
        | "left" -> LeftTab
        | "center" -> CenterTab
        | "right" -> RightTab
        | "decimal" -> DecimalTab
        | "bar" -> BarTab
        | other when other.StartsWith("other:") -> OtherTabAlignment(other.Substring(6))
        | other -> OtherTabAlignment other

    let private tabLeaderToStr (l: TabLeader) : string =
        match l with
        | NoLeader -> "none"
        | DotLeader -> "dot"
        | HyphenLeader -> "hyphen"
        | UnderscoreLeader -> "underscore"
        | HeavyLeader -> "heavy"
        | MiddleDotLeader -> "middleDot"

    let private tabLeaderOfStr (s: string) : TabLeader =
        match s with
        | "dot" -> DotLeader
        | "hyphen" -> HyphenLeader
        | "underscore" -> UnderscoreLeader
        | "heavy" -> HeavyLeader
        | "middleDot" -> MiddleDotLeader
        | _ -> NoLeader

    let private tabStopToJson (t: TabStop) : JsonNode =
        obj_
            [ "position", Some(jnum t.Position)
              "alignment", Some(jstr (tabAlignToStr t.Alignment))
              "leader", (if t.Leader = NoLeader then None else Some(jstr (tabLeaderToStr t.Leader))) ]
        :> JsonNode

    let private tabStopOfJson (n: JsonNode) : TabStop =
        let o = n.AsObject()

        { Position = num "position" o |> Option.defaultValue 0.0
          Alignment = str "alignment" o |> Option.map tabAlignOfStr |> Option.defaultValue LeftTab
          Leader = str "leader" o |> Option.map tabLeaderOfStr |> Option.defaultValue NoLeader }

    let private paragraphFormatToJson (f: ParagraphFormat) : JsonNode =
        obj_
            [ "alignment", f.Alignment |> Option.map (alignToStr >> jstr)
              "spacingBefore", f.SpacingBefore |> Option.map jnum
              "spacingAfter", f.SpacingAfter |> Option.map jnum
              "lineSpacing", f.LineSpacing |> Option.map lineSpacingToJson
              "indentation", f.Indentation |> Option.map indentationToJson
              "keepWithNext", (if f.KeepWithNext then Some(jbool true) else None)
              "pageBreakBefore", (if f.PageBreakBefore then Some(jbool true) else None)
              "shading", f.Shading |> Option.map (colorToStr >> jstr)
              "borders", f.Borders |> Option.map borderStyleToJson
              "tabStops", (if f.TabStops.IsEmpty then None else Some(JsonArray(f.TabStops |> List.map tabStopToJson |> Array.ofList) :> JsonNode)) ]
        :> JsonNode

    let private paragraphFormatOfJson (o: JsonObject) : ParagraphFormat =
        { Alignment = str "alignment" o |> Option.map alignOfStr
          SpacingBefore = num "spacingBefore" o
          SpacingAfter = num "spacingAfter" o
          LineSpacing = (match o.["lineSpacing"] with null -> None | n -> Some(lineSpacingOfJson n))
          Indentation = prop "indentation" o |> Option.map indentationOfJson
          KeepWithNext = boolean "keepWithNext" o
          PageBreakBefore = boolean "pageBreakBefore" o
          Shading = str "shading" o |> Option.map colorOfStr
          Borders = prop "borders" o |> Option.map borderStyleOfJson
          TabStops = arr "tabStops" o |> List.map tabStopOfJson }

    let private tableBordersToJson (b: TableBorders) : JsonNode =
        obj_
            [ "left", b.Outer.Left |> Option.map borderSideToJson
              "right", b.Outer.Right |> Option.map borderSideToJson
              "top", b.Outer.Top |> Option.map borderSideToJson
              "bottom", b.Outer.Bottom |> Option.map borderSideToJson
              "insideHorizontal", b.InsideHorizontal |> Option.map borderSideToJson
              "insideVertical", b.InsideVertical |> Option.map borderSideToJson ]
        :> JsonNode

    let private tableBordersOfJson (o: JsonObject) : TableBorders =
        { Outer =
            { Left = prop "left" o |> Option.map borderSideOfJson
              Right = prop "right" o |> Option.map borderSideOfJson
              Top = prop "top" o |> Option.map borderSideOfJson
              Bottom = prop "bottom" o |> Option.map borderSideOfJson }
          InsideHorizontal = prop "insideHorizontal" o |> Option.map borderSideOfJson
          InsideVertical = prop "insideVertical" o |> Option.map borderSideOfJson }

    // --- Images / hyperlinks -----------------------------------------------------------------

    let private imageToJson (img: ImageEntry) : JsonNode =
        obj_
            [ "format", Some(jstr (sprintf "%A" img.Format))
              "widthEmu", Some(jnum (float img.WidthEmu))
              "heightEmu", Some(jnum (float img.HeightEmu))
              "altText", img.AltText |> Option.map jstr
              "data", Some(jstr (Convert.ToBase64String(img.Data))) ]
        :> JsonNode

    let private imageOfJson (o: JsonObject) : ImageEntry =
        let format =
            match str "format" o with
            | Some "Jpeg" -> Jpeg
            | Some "Gif" -> Gif
            | Some "Bmp" -> Bmp
            | _ -> Png

        { Data = Convert.FromBase64String(str "data" o |> Option.get)
          Format = format
          WidthEmu = int64 (num "widthEmu" o |> Option.get)
          HeightEmu = int64 (num "heightEmu" o |> Option.get)
          AltText = str "altText" o }

    let private hyperlinkTargetToJson (t: HyperlinkTarget) : JsonNode =
        match t with
        | ExternalUrl u -> obj_ [ "externalHyperlink", Some(jstr u) ] :> JsonNode
        | InternalBookmark n -> obj_ [ "internalHyperlink", Some(jstr n) ] :> JsonNode

    let private hyperlinkTargetOfJson (o: JsonObject) : HyperlinkTarget =
        match str "externalHyperlink" o, str "internalHyperlink" o with
        | Some u, _ -> ExternalUrl u
        | _, Some n -> InternalBookmark n
        | _ -> ExternalUrl ""

    // --- Inline content ---------------------------------------------------------------------

    // --- Paragraphs / tables -----------------------------------------------------------------

    let private tableStyleRefToJson (s: TableStyleRef) : JsonNode =
        obj_
            [ "name", Some(jstr s.Name)
              "firstRow", (if s.FirstRowBanding then Some(jbool true) else None)
              "lastRow", (if s.LastRowBanding then Some(jbool true) else None)
              "bandedRows", (if s.BandedRows then Some(jbool true) else None)
              "bandedColumns", (if s.BandedColumns then Some(jbool true) else None) ]
        :> JsonNode

    let private tableStyleRefOfJson (o: JsonObject) : TableStyleRef =
        { Name = str "name" o |> Option.get
          FirstRowBanding = boolean "firstRow" o
          LastRowBanding = boolean "lastRow" o
          BandedRows = boolean "bandedRows" o
          BandedColumns = boolean "bandedColumns" o }

    let private cellMarginsToJson (m: CellMargins) : JsonNode =
        obj_ [ "top", m.Top |> Option.map jnum; "bottom", m.Bottom |> Option.map jnum; "left", m.Left |> Option.map jnum; "right", m.Right |> Option.map jnum ] :> JsonNode

    let private cellMarginsOfJson (o: JsonObject) : CellMargins =
        { Top = num "top" o; Bottom = num "bottom" o; Left = num "left" o; Right = num "right" o }

    let private tableCellPropsToJson (p: TableCellProps) : JsonNode =
        obj_
            [ "gridSpan", p.GridSpan |> Option.map jint
              "verticalMerge", p.VerticalMerge |> Option.map (function RestartMerge -> jstr "restart" | ContinueMerge -> jstr "continue")
              "shading", p.Shading |> Option.map (colorToStr >> jstr)
              "borders", p.Borders |> Option.map tableBordersToJson
              "width", p.Width |> Option.map jnum ]
        :> JsonNode

    let private tableCellPropsOfJson (o: JsonObject) : TableCellProps =
        { GridSpan = intg "gridSpan" o
          VerticalMerge = str "verticalMerge" o |> Option.map (fun s -> if s = "continue" then ContinueMerge else RestartMerge)
          Shading = str "shading" o |> Option.map colorOfStr
          Borders = prop "borders" o |> Option.map tableBordersOfJson
          Width = num "width" o }

    // `inlineToJson`/`inlineOfJson` need `blockToJson`/`blockOfJson` (a `Footnote`/
    // `Endnote`'s own body is a `Block list`), which need `paragraphToJson`/
    // `paragraphOfJson`, which need `inlineToJson`/`inlineOfJson` back for a paragraph's
    // own `Inlines` - one `rec ... and ...` chain, same cycle `Xml.fs`'s equivalent
    // functions are chained for.
    let rec private inlineToJson (i: Inline) : JsonNode =
        match i with
        | Run(text, style, styleId) ->
            obj_ [ "run", Some(obj_ [ "text", Some(jstr text); "style", style |> Option.map runStyleToJson; "styleId", styleId |> Option.map jstr ] :> JsonNode) ] :> JsonNode
        | LineBreak -> jstr "lineBreak"
        | Tab -> jstr "tab"
        | PageBreak -> jstr "pageBreak"
        | Image img -> obj_ [ "image", Some(imageToJson img) ] :> JsonNode
        | Hyperlink(target, runs, tooltip) ->
            obj_
                [ "hyperlink",
                  Some(
                      obj_
                          [ "target", Some(hyperlinkTargetToJson target)
                            "runs", Some(JsonArray(runs |> List.map inlineToJson |> Array.ofList))
                            "tooltip", tooltip |> Option.map jstr ]
                      :> JsonNode
                  ) ]
            :> JsonNode
        | Bookmark(name, content) ->
            obj_ [ "bookmark", Some(obj_ [ "name", Some(jstr name); "content", Some(JsonArray(content |> List.map inlineToJson |> Array.ofList)) ] :> JsonNode) ] :> JsonNode
        | Comment(author, initials, date, text, content) ->
            obj_
                [ "comment",
                  Some(
                      obj_
                          [ "author", Some(jstr author)
                            "initials", initials |> Option.map jstr
                            "date", date |> Option.map (fun d -> jstr (d.ToString("o")))
                            "text", Some(jstr text)
                            "content", Some(JsonArray(content |> List.map inlineToJson |> Array.ofList)) ]
                      :> JsonNode
                  ) ]
            :> JsonNode
        | Field(instr, cached) ->
            obj_ [ "field", Some(obj_ [ "instruction", Some(jstr instr); "cachedResult", cached |> Option.map jstr ] :> JsonNode) ] :> JsonNode
        | Footnote content -> obj_ [ "footnote", Some(JsonArray(content |> List.map blockToJson |> Array.ofList) :> JsonNode) ] :> JsonNode
        | Endnote content -> obj_ [ "endnote", Some(JsonArray(content |> List.map blockToJson |> Array.ofList) :> JsonNode) ] :> JsonNode

    and private inlineOfJson (n: JsonNode) : Inline =
        match n with
        | :? JsonValue as v ->
            match v.GetValue<string>() with
            | "lineBreak" -> LineBreak
            | "tab" -> Tab
            | _ -> PageBreak
        | _ ->
            let o = n.AsObject()

            if not (isNull o.["run"]) then
                let r = o.["run"].AsObject()
                Run(str "text" r |> Option.get, prop "style" r |> Option.map runStyleOfJson, str "styleId" r)
            elif not (isNull o.["image"]) then
                Image(imageOfJson (o.["image"].AsObject()))
            elif not (isNull o.["hyperlink"]) then
                let h = o.["hyperlink"].AsObject()
                let target = hyperlinkTargetOfJson (h.["target"].AsObject())
                let runs = arr "runs" h |> List.map inlineOfJson
                Hyperlink(target, runs, str "tooltip" h)
            elif not (isNull o.["bookmark"]) then
                let b = o.["bookmark"].AsObject()
                Bookmark(str "name" b |> Option.get, arr "content" b |> List.map inlineOfJson)
            elif not (isNull o.["comment"]) then
                let c = o.["comment"].AsObject()
                Comment(str "author" c |> Option.get, str "initials" c, str "date" c |> Option.map DateTime.Parse, str "text" c |> Option.get, arr "content" c |> List.map inlineOfJson)
            elif not (isNull o.["field"]) then
                let f = o.["field"].AsObject()
                Field(str "instruction" f |> Option.get, str "cachedResult" f)
            elif not (isNull o.["footnote"]) then
                Footnote(o.["footnote"].AsArray() |> Seq.map blockOfJson |> List.ofSeq)
            elif not (isNull o.["endnote"]) then
                Endnote(o.["endnote"].AsArray() |> Seq.map blockOfJson |> List.ofSeq)
            else
                failwith "Unknown inline JSON shape"

    and private paragraphToJson (p: Paragraph) : JsonNode =
        obj_
            [ "styleId", p.StyleId |> Option.map jstr
              "format", p.Format |> Option.map paragraphFormatToJson
              "numId", p.Numbering |> Option.map (fst >> jint)
              "level", p.Numbering |> Option.map (snd >> jint)
              "inlines", Some(JsonArray(p.Inlines |> List.map inlineToJson |> Array.ofList)) ]
        :> JsonNode

    and private paragraphOfJson (o: JsonObject) : Paragraph =
        { Inlines = arr "inlines" o |> List.map inlineOfJson
          StyleId = str "styleId" o
          Format = prop "format" o |> Option.map paragraphFormatOfJson
          Numbering = match intg "numId" o, intg "level" o with
                      | Some n, Some l -> Some(n, l)
                      | _ -> None }

    and private blockToJson (b: Block) : JsonNode =
        match b with
        | ParagraphBlock p -> obj_ [ "para", Some(paragraphToJson p) ] :> JsonNode
        | TableBlock t -> obj_ [ "table", Some(tableToJson t) ] :> JsonNode

    and private blockOfJson (n: JsonNode) : Block =
        let o = n.AsObject()

        if not (isNull o.["para"]) then
            ParagraphBlock(paragraphOfJson (o.["para"].AsObject()))
        else
            TableBlock(tableOfJson (o.["table"].AsObject()))

    and private tableCellToJson (c: TableCell) : JsonNode =
        obj_ [ "props", Some(tableCellPropsToJson c.Props); "content", Some(JsonArray(c.Content |> List.map blockToJson |> Array.ofList)) ] :> JsonNode

    and private tableCellOfJson (n: JsonNode) : TableCell =
        let o = n.AsObject()
        { Content = arr "content" o |> List.map blockOfJson
          Props = prop "props" o |> Option.map tableCellPropsOfJson |> Option.defaultValue TableCellProps.Default }

    and private tableRowToJson (r: TableRow) : JsonNode =
        obj_
            [ "height", r.Height |> Option.map jnum
              "repeatAsHeader", (if r.RepeatAsHeader then Some(jbool true) else None)
              "cells", Some(JsonArray(r.Cells |> List.map tableCellToJson |> Array.ofList)) ]
        :> JsonNode

    and private tableRowOfJson (n: JsonNode) : TableRow =
        let o = n.AsObject()
        { Cells = arr "cells" o |> List.map tableCellOfJson
          Height = num "height" o
          RepeatAsHeader = boolean "repeatAsHeader" o }

    and private tableToJson (t: TableEntry) : JsonNode =
        obj_
            [ "style", t.Style |> Option.map tableStyleRefToJson
              "borders", t.Borders |> Option.map tableBordersToJson
              "cellMargins", t.CellMargins |> Option.map cellMarginsToJson
              "columnWidths", Some(JsonArray(t.ColumnWidths |> List.map jnum |> Array.ofList))
              "rows", Some(JsonArray(t.Rows |> List.map tableRowToJson |> Array.ofList)) ]
        :> JsonNode

    and private tableOfJson (o: JsonObject) : TableEntry =
        { Rows = arr "rows" o |> List.map tableRowOfJson
          ColumnWidths = arr "columnWidths" o |> List.map (fun n -> n.GetValue<float>())
          Style = prop "style" o |> Option.map tableStyleRefOfJson
          Borders = prop "borders" o |> Option.map tableBordersOfJson
          CellMargins = prop "cellMargins" o |> Option.map cellMarginsOfJson }

    // --- Page setup / headers & footers -------------------------------------------------------

    let private pageSizeToJson (p: PageSize) : JsonNode =
        match p with
        | OtherPageSize code -> obj_ [ "other", Some(jint code) ] :> JsonNode
        | CustomPageSize(w, h) -> obj_ [ "widthPoints", Some(jnum w); "heightPoints", Some(jnum h) ] :> JsonNode
        | named -> jstr (sprintf "%A" named)

    let private pageSizeOfJson (n: JsonNode) : PageSize =
        match n with
        | :? JsonValue as v ->
            match v.GetValue<string>() with
            | "Legal" -> Legal
            | "A4" -> A4
            | "A3" -> A3
            | _ -> Letter
        | _ ->
            let o = n.AsObject()

            match intg "other" o, num "widthPoints" o, num "heightPoints" o with
            | Some code, _, _ -> OtherPageSize code
            | _, Some w, Some h -> CustomPageSize(w, h)
            | _ -> Letter

    let private pageMarginsToJson (m: PageMargins) : JsonNode =
        obj_
            [ "top", Some(jnum m.Top)
              "bottom", Some(jnum m.Bottom)
              "left", Some(jnum m.Left)
              "right", Some(jnum m.Right)
              "header", Some(jnum m.Header)
              "footer", Some(jnum m.Footer)
              "gutter", Some(jnum m.Gutter) ]
        :> JsonNode

    let private pageMarginsOfJson (o: JsonObject) : PageMargins =
        { Top = num "top" o |> Option.get
          Bottom = num "bottom" o |> Option.get
          Left = num "left" o |> Option.get
          Right = num "right" o |> Option.get
          Header = num "header" o |> Option.get
          Footer = num "footer" o |> Option.get
          Gutter = num "gutter" o |> Option.get }

    let private headerFooterSetToJson (h: HeaderFooterSet) : JsonNode =
        obj_
            [ "default", h.Default |> Option.map (fun bs -> JsonArray(bs |> List.map blockToJson |> Array.ofList) :> JsonNode)
              "first", h.First |> Option.map (fun bs -> JsonArray(bs |> List.map blockToJson |> Array.ofList) :> JsonNode)
              "even", h.Even |> Option.map (fun bs -> JsonArray(bs |> List.map blockToJson |> Array.ofList) :> JsonNode) ]
        :> JsonNode

    let private headerFooterSetOfJson (o: JsonObject) : HeaderFooterSet =
        let variant (name: string) = match o.[name] with null -> None | n -> Some(n.AsArray() |> Seq.map blockOfJson |> List.ofSeq)
        { Default = variant "default"; First = variant "first"; Even = variant "even" }

    let private sectionPropertiesToJson (s: SectionProperties) : JsonNode =
        obj_
            [ "pageSize", Some(pageSizeToJson s.PageSize)
              "orientation", Some(jstr (sprintf "%A" s.Orientation))
              "margins", Some(pageMarginsToJson s.Margins)
              "header", s.Header |> Option.map headerFooterSetToJson
              "footer", s.Footer |> Option.map headerFooterSetToJson
              "pageNumberStart", s.PageNumberStart |> Option.map jint
              "columns", Some(jint s.Columns)
              "breakType", Some(jstr (sprintf "%A" s.BreakType)) ]
        :> JsonNode

    let private sectionBreakTypeOfJson (s: string option) : SectionBreakType =
        match s with
        | Some "ContinuousBreak" -> ContinuousBreak
        | Some "EvenPageBreak" -> EvenPageBreak
        | Some "OddPageBreak" -> OddPageBreak
        | _ -> NextPageBreak

    let private sectionPropertiesOfJson (o: JsonObject) : SectionProperties =
        { PageSize = pageSizeOfJson o.["pageSize"]
          Orientation = (if str "orientation" o = Some "Landscape" then Landscape else Portrait)
          Margins = pageMarginsOfJson (prop "margins" o |> Option.get)
          Header = prop "header" o |> Option.map headerFooterSetOfJson
          Footer = prop "footer" o |> Option.map headerFooterSetOfJson
          PageNumberStart = intg "pageNumberStart" o
          Columns = intg "columns" o |> Option.defaultValue 1
          BreakType = sectionBreakTypeOfJson (str "breakType" o) }

    let private sectionToJson (s: Section) : JsonNode =
        obj_ [ "pageSetup", Some(sectionPropertiesToJson s.Properties); "body", Some(JsonArray(s.Body |> List.map blockToJson |> Array.ofList)) ] :> JsonNode

    let private sectionOfJson (n: JsonNode) : Section =
        let o = n.AsObject()
        { Body = arr "body" o |> List.map blockOfJson; Properties = sectionPropertiesOfJson (prop "pageSetup" o |> Option.get) }

    // --- Styles / numbering / protection -----------------------------------------------------

    let private styleDefinitionToJson (d: StyleDefinition) : JsonNode =
        obj_
            [ "id", Some(jstr d.Id)
              "name", Some(jstr d.Name)
              "type", Some(jstr (match d.Type with ParagraphStyleType -> "paragraph" | CharacterStyleType -> "character"))
              "basedOn", d.BasedOn |> Option.map jstr
              "runStyle", d.RunFormat |> Option.map runStyleToJson
              "paragraphFormat", d.ParaFormat |> Option.map paragraphFormatToJson ]
        :> JsonNode

    let private styleDefinitionOfJson (n: JsonNode) : StyleDefinition =
        let o = n.AsObject()
        { Id = str "id" o |> Option.get
          Name = str "name" o |> Option.get
          Type = (if str "type" o = Some "character" then CharacterStyleType else ParagraphStyleType)
          BasedOn = str "basedOn" o
          RunFormat = prop "runStyle" o |> Option.map runStyleOfJson
          ParaFormat = prop "paragraphFormat" o |> Option.map paragraphFormatOfJson }

    let private numberFormatKindToJson (k: NumberFormatKind) : JsonNode =
        match k with
        | BulletFormat(glyph, font) -> obj_ [ "bullet", Some(obj_ [ "glyph", Some(jint (int glyph)); "font", Some(jstr font) ] :> JsonNode) ] :> JsonNode
        | OtherFormat raw -> obj_ [ "other", Some(jstr raw) ] :> JsonNode
        | other -> jstr (sprintf "%A" other)

    let private numberFormatKindOfJson (n: JsonNode) : NumberFormatKind =
        match n with
        | :? JsonValue as v ->
            match v.GetValue<string>() with
            | "LowerLetterFormat" -> LowerLetterFormat
            | "UpperLetterFormat" -> UpperLetterFormat
            | "LowerRomanFormat" -> LowerRomanFormat
            | "UpperRomanFormat" -> UpperRomanFormat
            | _ -> DecimalFormat
        | _ ->
            let o = n.AsObject()

            if not (isNull o.["bullet"]) then
                let b = o.["bullet"].AsObject()
                BulletFormat(char (intg "glyph" b |> Option.get), str "font" b |> Option.get)
            elif not (isNull o.["other"]) then
                OtherFormat(str "other" o |> Option.get)
            else
                DecimalFormat

    let private listLevelToJson (l: ListLevel) : JsonNode =
        obj_
            [ "format", Some(numberFormatKindToJson l.Format)
              "text", Some(jstr l.Text)
              "indentLeft", l.IndentLeft |> Option.map jnum
              "hangingIndent", l.HangingIndent |> Option.map jnum
              "startAt", l.StartAt |> Option.map jint ]
        :> JsonNode

    let private listLevelOfJson (n: JsonNode) : ListLevel =
        let o = n.AsObject()
        { Format = numberFormatKindOfJson o.["format"]
          Text = str "text" o |> Option.get
          IndentLeft = num "indentLeft" o
          HangingIndent = num "hangingIndent" o
          StartAt = intg "startAt" o }

    let private numberingDefinitionToJson (d: NumberingDefinition) : JsonNode =
        obj_ [ "id", Some(jint d.Id); "levels", Some(JsonArray(d.Levels |> List.map listLevelToJson |> Array.ofList)) ] :> JsonNode

    let private numberingDefinitionOfJson (n: JsonNode) : NumberingDefinition =
        let o = n.AsObject()
        { Id = intg "id" o |> Option.get; Levels = arr "levels" o |> List.map listLevelOfJson }

    let private protectionToJson (p: DocumentProtection) : JsonNode =
        obj_ [ "edit", p.Edit |> Option.map (sprintf "%A" >> jstr); "password", p.Password |> Option.map jstr ] :> JsonNode

    let private protectionOfJson (o: JsonObject) : DocumentProtection =
        let edit =
            str "edit" o
            |> Option.map (function
                | "CommentsOnlyRestriction" -> CommentsOnlyRestriction
                | "TrackedChangesOnlyRestriction" -> TrackedChangesOnlyRestriction
                | "FormsOnlyRestriction" -> FormsOnlyRestriction
                | _ -> ReadOnlyRestriction)

        { Edit = edit; Password = str "password" o }

    let private documentPropertiesToJson (p: DocumentProperties) : JsonNode =
        obj_
            [ "title", p.Title |> Option.map jstr
              "author", p.Author |> Option.map jstr
              "subject", p.Subject |> Option.map jstr
              "keywords", p.Keywords |> Option.map jstr
              "comments", p.Comments |> Option.map jstr
              "category", p.Category |> Option.map jstr
              "company", p.Company |> Option.map jstr ]
        :> JsonNode

    let private documentPropertiesOfJson (o: JsonObject) : DocumentProperties =
        { Title = str "title" o
          Author = str "author" o
          Subject = str "subject" o
          Keywords = str "keywords" o
          Comments = str "comments" o
          Category = str "category" o
          Company = str "company" o }

    let private tableStyleRegionToJson (r: TableStyleRegion) : JsonNode option =
        if r = TableStyleRegion.None then
            None
        else
            Some(
                obj_
                    [ "runStyle", r.RunFormat |> Option.map runStyleToJson
                      "paragraphFormat", r.ParaFormat |> Option.map paragraphFormatToJson
                      "cellShading", r.CellShading |> Option.map (colorToStr >> jstr) ]
                :> JsonNode
            )

    let private tableStyleRegionOfJson (o: JsonObject) : TableStyleRegion =
        { RunFormat = prop "runStyle" o |> Option.map runStyleOfJson
          ParaFormat = prop "paragraphFormat" o |> Option.map paragraphFormatOfJson
          CellShading = str "cellShading" o |> Option.map colorOfStr }

    let private tableStyleDefinitionToJson (d: TableStyleDefinition) : JsonNode =
        obj_
            [ "id", Some(jstr d.Id)
              "name", Some(jstr d.Name)
              "basedOn", d.BasedOn |> Option.map jstr
              "borders", d.Borders |> Option.map tableBordersToJson
              "wholeTable", tableStyleRegionToJson d.WholeTable
              "firstRow", tableStyleRegionToJson d.FirstRow
              "bandedRow", tableStyleRegionToJson d.BandedRow ]
        :> JsonNode

    let private tableStyleDefinitionOfJson (n: JsonNode) : TableStyleDefinition =
        let o = n.AsObject()

        { Id = str "id" o |> Option.get
          Name = str "name" o |> Option.get
          BasedOn = str "basedOn" o
          Borders = prop "borders" o |> Option.map tableBordersOfJson
          WholeTable = prop "wholeTable" o |> Option.map tableStyleRegionOfJson |> Option.defaultValue TableStyleRegion.None
          FirstRow = prop "firstRow" o |> Option.map tableStyleRegionOfJson |> Option.defaultValue TableStyleRegion.None
          BandedRow = prop "bandedRow" o |> Option.map tableStyleRegionOfJson |> Option.defaultValue TableStyleRegion.None }

    // --- Top level ------------------------------------------------------------------------

    /// `Document` -> `JsonObject`. See this file's own conventions above and the worked
    /// examples in the root README.
    let toDocument (doc: Document) : JsonObject =
        obj_
            [ "sections", Some(JsonArray(doc.Sections |> List.map sectionToJson |> Array.ofList))
              "styles", (if doc.Styles.IsEmpty then None else Some(JsonArray(doc.Styles |> List.map styleDefinitionToJson |> Array.ofList)))
              "numbering", (if doc.Numbering.IsEmpty then None else Some(JsonArray(doc.Numbering |> List.map numberingDefinitionToJson |> Array.ofList)))
              "protection", doc.Protection |> Option.map protectionToJson
              "vbaProject", doc.VbaProject |> Option.map (fun b -> jstr (Convert.ToBase64String(b)))
              "properties", (if doc.Properties = DocumentProperties.Default then None else Some(documentPropertiesToJson doc.Properties))
              "tableStyles", (if doc.TableStyles.IsEmpty then None else Some(JsonArray(doc.TableStyles |> List.map tableStyleDefinitionToJson |> Array.ofList))) ]

    /// `JsonObject` -> `Document`, the inverse of `toDocument`.
    let ofDocument (o: JsonObject) : Document =
        { Sections = arr "sections" o |> List.map sectionOfJson
          Styles = arr "styles" o |> List.map styleDefinitionOfJson
          Numbering = arr "numbering" o |> List.map numberingDefinitionOfJson
          Protection = prop "protection" o |> Option.map protectionOfJson
          VbaProject = str "vbaProject" o |> Option.map Convert.FromBase64String
          Properties = prop "properties" o |> Option.map documentPropertiesOfJson |> Option.defaultValue DocumentProperties.Default
          TableStyles = arr "tableStyles" o |> List.map tableStyleDefinitionOfJson }
