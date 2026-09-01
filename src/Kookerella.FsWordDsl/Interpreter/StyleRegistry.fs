namespace Kookerella.FsWordDsl.Interpreter

open System
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open Kookerella.FsWordDsl

/// Formatting conversions shared by `Writer`/`Reader`: `Styles.Color`/`RunStyle`/
/// `ParagraphFormat`/`BorderStyle` <-> the equivalent `DocumentFormat.OpenXml.Wordprocessing`
/// element shapes, plus `Document.Styles` <-> `styles.xml`. Unlike Excel's own
/// `StyleRegistry` (which interns `CellStyle` values into a shared, indexed stylesheet -
/// direct formatting has no such index in WordprocessingML, it's written inline on every
/// run/paragraph's own `w:rPr`/`w:pPr`), this module's only "registry" responsibility is
/// ensuring every referenced named `StyleId` actually has a real `<w:style>` entry, and that
/// `"Normal"` always exists - see `stylesToOpenXml`.
module StyleRegistry =

    // --- Color / underline / highlight -------------------------------------------------

    let colorToHex (c: Color) : string =
        match c with
        | Rgb(r, g, b) -> sprintf "%02X%02X%02X" r g b
        | Auto -> "auto"

    let colorOfHex (hex: string) : Color =
        if String.Equals(hex, "auto", StringComparison.OrdinalIgnoreCase) then
            Auto
        else
            let n = Convert.ToInt32(hex, 16)
            Rgb(byte ((n >>> 16) &&& 0xFF), byte ((n >>> 8) &&& 0xFF), byte (n &&& 0xFF))

    let underlineToW (u: UnderlineStyle) : Wordprocessing.UnderlineValues =
        match u with
        | SingleUnderline -> Wordprocessing.UnderlineValues.Single
        | DoubleUnderline -> Wordprocessing.UnderlineValues.Double
        | ThickUnderline -> Wordprocessing.UnderlineValues.Thick
        | DottedUnderline -> Wordprocessing.UnderlineValues.Dotted
        | DashedUnderline -> Wordprocessing.UnderlineValues.Dash
        | WavyUnderline -> Wordprocessing.UnderlineValues.Wave
        | OtherUnderline raw -> Wordprocessing.UnderlineValues raw

    let underlineOfW (v: Wordprocessing.UnderlineValues) : UnderlineStyle =
        if v = Wordprocessing.UnderlineValues.Single then SingleUnderline
        elif v = Wordprocessing.UnderlineValues.Double then DoubleUnderline
        elif v = Wordprocessing.UnderlineValues.Thick then ThickUnderline
        elif v = Wordprocessing.UnderlineValues.Dotted then DottedUnderline
        elif v = Wordprocessing.UnderlineValues.Dash then DashedUnderline
        elif v = Wordprocessing.UnderlineValues.Wave then WavyUnderline
        else OtherUnderline(v.ToString())

    let highlightToW (h: HighlightColor) : Wordprocessing.HighlightColorValues =
        match h with
        | HlYellow -> Wordprocessing.HighlightColorValues.Yellow
        | HlGreen -> Wordprocessing.HighlightColorValues.Green
        | HlCyan -> Wordprocessing.HighlightColorValues.Cyan
        | HlMagenta -> Wordprocessing.HighlightColorValues.Magenta
        | HlBlue -> Wordprocessing.HighlightColorValues.Blue
        | HlRed -> Wordprocessing.HighlightColorValues.Red
        | HlDarkBlue -> Wordprocessing.HighlightColorValues.DarkBlue
        | HlDarkCyan -> Wordprocessing.HighlightColorValues.DarkCyan
        | HlDarkGreen -> Wordprocessing.HighlightColorValues.DarkGreen
        | HlDarkMagenta -> Wordprocessing.HighlightColorValues.DarkMagenta
        | HlDarkRed -> Wordprocessing.HighlightColorValues.DarkRed
        | HlDarkYellow -> Wordprocessing.HighlightColorValues.DarkYellow
        | HlDarkGray -> Wordprocessing.HighlightColorValues.DarkGray
        | HlLightGray -> Wordprocessing.HighlightColorValues.LightGray
        | HlBlack -> Wordprocessing.HighlightColorValues.Black

    let highlightOfW (v: Wordprocessing.HighlightColorValues) : HighlightColor option =
        if v = Wordprocessing.HighlightColorValues.Yellow then Some HlYellow
        elif v = Wordprocessing.HighlightColorValues.Green then Some HlGreen
        elif v = Wordprocessing.HighlightColorValues.Cyan then Some HlCyan
        elif v = Wordprocessing.HighlightColorValues.Magenta then Some HlMagenta
        elif v = Wordprocessing.HighlightColorValues.Blue then Some HlBlue
        elif v = Wordprocessing.HighlightColorValues.Red then Some HlRed
        elif v = Wordprocessing.HighlightColorValues.DarkBlue then Some HlDarkBlue
        elif v = Wordprocessing.HighlightColorValues.DarkCyan then Some HlDarkCyan
        elif v = Wordprocessing.HighlightColorValues.DarkGreen then Some HlDarkGreen
        elif v = Wordprocessing.HighlightColorValues.DarkMagenta then Some HlDarkMagenta
        elif v = Wordprocessing.HighlightColorValues.DarkRed then Some HlDarkRed
        elif v = Wordprocessing.HighlightColorValues.DarkYellow then Some HlDarkYellow
        elif v = Wordprocessing.HighlightColorValues.DarkGray then Some HlDarkGray
        elif v = Wordprocessing.HighlightColorValues.LightGray then Some HlLightGray
        elif v = Wordprocessing.HighlightColorValues.Black then Some HlBlack
        else None

    // --- Borders -------------------------------------------------------------------------

    let borderLineStyleToW (s: BorderLineStyle) : Wordprocessing.BorderValues =
        match s with
        | SingleLine -> Wordprocessing.BorderValues.Single
        | ThickLine -> Wordprocessing.BorderValues.Thick
        | DoubleLine -> Wordprocessing.BorderValues.Double
        | DottedLine -> Wordprocessing.BorderValues.Dotted
        | DashedLine -> Wordprocessing.BorderValues.Dashed
        | WaveLine -> Wordprocessing.BorderValues.Wave
        | OtherLine raw -> Wordprocessing.BorderValues raw

    let borderLineStyleOfW (v: Wordprocessing.BorderValues) : BorderLineStyle =
        if v = Wordprocessing.BorderValues.Single then SingleLine
        elif v = Wordprocessing.BorderValues.Thick then ThickLine
        elif v = Wordprocessing.BorderValues.Double then DoubleLine
        elif v = Wordprocessing.BorderValues.Dotted then DottedLine
        elif v = Wordprocessing.BorderValues.Dashed then DashedLine
        elif v = Wordprocessing.BorderValues.Wave then WaveLine
        else OtherLine(v.ToString())

    /// Border `sz` is in eighths of a point; `None` writes OOXML's own default weight (4 =
    /// half a point).
    let borderSideWidthEighths (side: BorderSide) : uint32 =
        match side.Width with
        | Some pts -> uint32 (Math.Round(pts * 8.0))
        | None -> 4u

    let borderSideOfWidthEighths (eighths: uint32) : float = float eighths / 8.0

    // --- Run / paragraph properties --------------------------------------------------------

    let runPropertiesOf (style: RunStyle option) (styleId: string option) : Wordprocessing.RunProperties option =
        if style.IsNone && styleId.IsNone then
            None
        else
            let rPr = Wordprocessing.RunProperties()
            styleId |> Option.iter (fun id -> rPr.RunStyle <- Wordprocessing.RunStyle(Val = StringValue id))

            style
            |> Option.iter (fun s ->
                s.FontFamily
                |> Option.iter (fun f ->
                    rPr.RunFonts <- Wordprocessing.RunFonts(Ascii = StringValue f, HighAnsi = StringValue f))

                s.Size
                |> Option.iter (fun sz ->
                    let halfPoints = StringValue(string (int (Math.Round(sz * 2.0))))
                    rPr.FontSize <- Wordprocessing.FontSize(Val = halfPoints)
                    rPr.FontSizeComplexScript <- Wordprocessing.FontSizeComplexScript(Val = halfPoints))

                if s.Bold then
                    rPr.Bold <- Wordprocessing.Bold()

                if s.Italic then
                    rPr.Italic <- Wordprocessing.Italic()

                s.Underline
                |> Option.iter (fun u -> rPr.Underline <- Wordprocessing.Underline(Val = EnumValue(underlineToW u)))

                if s.Strikethrough then
                    rPr.Strike <- Wordprocessing.Strike()

                s.Color
                |> Option.iter (fun c -> rPr.Color <- Wordprocessing.Color(Val = StringValue(colorToHex c)))

                s.Highlight
                |> Option.iter (fun h -> rPr.Highlight <- Wordprocessing.Highlight(Val = EnumValue(highlightToW h)))

                s.VerticalPosition
                |> Option.iter (fun vp ->
                    let v =
                        match vp with
                        | Superscript -> Wordprocessing.VerticalPositionValues.Superscript
                        | Subscript -> Wordprocessing.VerticalPositionValues.Subscript

                    rPr.VerticalTextAlignment <- Wordprocessing.VerticalTextAlignment(Val = EnumValue v)))

            Some rPr

    /// The inverse of `runPropertiesOf` - `styleId` is read separately from `RunStyle` since
    /// they're independent DSL fields (see `Model.Inline.Run`).
    let runStyleOfProperties (rPr: Wordprocessing.RunProperties option) : RunStyle option =
        match rPr with
        | None -> None
        | Some rPr ->
            let hasAny =
                isNull rPr.RunFonts |> not
                || isNull rPr.FontSize |> not
                || isNull rPr.Bold |> not
                || isNull rPr.Italic |> not
                || isNull rPr.Underline |> not
                || isNull rPr.Strike |> not
                || isNull rPr.Color |> not
                || isNull rPr.Highlight |> not
                || isNull rPr.VerticalTextAlignment |> not

            if not hasAny then
                None
            else
                Some
                    { RunStyle.Default with
                        FontFamily =
                            if isNull rPr.RunFonts then None
                            else rPr.RunFonts.Ascii |> Option.ofObj |> Option.map (fun v -> v.Value)
                        Size =
                            if isNull rPr.FontSize then None
                            else rPr.FontSize.Val |> Option.ofObj |> Option.map (fun v -> float v.Value / 2.0)
                        Bold = not (isNull rPr.Bold)
                        Italic = not (isNull rPr.Italic)
                        Underline =
                            if isNull rPr.Underline then None
                            else rPr.Underline.Val |> Option.ofObj |> Option.map (fun v -> underlineOfW v.Value)
                        Strikethrough = not (isNull rPr.Strike)
                        Color =
                            if isNull rPr.Color then None
                            else rPr.Color.Val |> Option.ofObj |> Option.map (fun v -> colorOfHex v.Value)
                        Highlight =
                            if isNull rPr.Highlight then None
                            else rPr.Highlight.Val |> Option.ofObj |> Option.bind (fun v -> highlightOfW v.Value)
                        VerticalPosition =
                            if isNull rPr.VerticalTextAlignment then None
                            else
                                rPr.VerticalTextAlignment.Val
                                |> Option.ofObj
                                |> Option.map (fun v ->
                                    if v.Value = Wordprocessing.VerticalPositionValues.Superscript then Superscript else Subscript) }

    let styleIdOfProperties (rPr: Wordprocessing.RunProperties option) : string option =
        match rPr with
        | Some rPr when not (isNull rPr.RunStyle) -> rPr.RunStyle.Val |> Option.ofObj |> Option.map (fun v -> v.Value)
        | _ -> None

    let private twentiethsOfPoint (pts: float) : int = int (Math.Round(pts * 20.0))
    let private pointsOfTwentieths (v: int) : float = float v / 20.0

    let indentationToW (ind: Indentation) : Wordprocessing.Indentation =
        let i = Wordprocessing.Indentation()
        ind.Left |> Option.iter (fun v -> i.Left <- StringValue(string (twentiethsOfPoint v)))
        ind.Right |> Option.iter (fun v -> i.Right <- StringValue(string (twentiethsOfPoint v)))
        ind.FirstLine |> Option.iter (fun v -> i.FirstLine <- StringValue(string (twentiethsOfPoint v)))
        ind.Hanging |> Option.iter (fun v -> i.Hanging <- StringValue(string (twentiethsOfPoint v)))
        i

    let indentationOfW (i: Wordprocessing.Indentation) : Indentation =
        { Left = i.Left |> Option.ofObj |> Option.map (fun v -> pointsOfTwentieths (int v.Value))
          Right = i.Right |> Option.ofObj |> Option.map (fun v -> pointsOfTwentieths (int v.Value))
          FirstLine = i.FirstLine |> Option.ofObj |> Option.map (fun v -> pointsOfTwentieths (int v.Value))
          Hanging = i.Hanging |> Option.ofObj |> Option.map (fun v -> pointsOfTwentieths (int v.Value)) }

    let paragraphPropertiesOf (styleId: string option) (format: ParagraphFormat option) : Wordprocessing.ParagraphProperties option =
        if styleId.IsNone && format.IsNone then
            None
        else
            let pPr = Wordprocessing.ParagraphProperties()
            styleId |> Option.iter (fun id -> pPr.ParagraphStyleId <- Wordprocessing.ParagraphStyleId(Val = StringValue id))

            format
            |> Option.iter (fun f ->
                f.Alignment
                |> Option.iter (fun a ->
                    let v =
                        match a with
                        | AlignLeft -> Wordprocessing.JustificationValues.Left
                        | AlignCenter -> Wordprocessing.JustificationValues.Center
                        | AlignRight -> Wordprocessing.JustificationValues.Right
                        | AlignJustify -> Wordprocessing.JustificationValues.Both

                    pPr.Justification <- Wordprocessing.Justification(Val = EnumValue v))

                if f.SpacingBefore.IsSome || f.SpacingAfter.IsSome || f.LineSpacing.IsSome then
                    let spacing = Wordprocessing.SpacingBetweenLines()
                    f.SpacingBefore |> Option.iter (fun v -> spacing.Before <- StringValue(string (twentiethsOfPoint v)))
                    f.SpacingAfter |> Option.iter (fun v -> spacing.After <- StringValue(string (twentiethsOfPoint v)))

                    f.LineSpacing
                    |> Option.iter (fun ls ->
                        match ls with
                        | SingleSpacing ->
                            spacing.Line <- StringValue "240"
                            spacing.LineRule <- EnumValue Wordprocessing.LineSpacingRuleValues.Auto
                        | OnePointFiveSpacing ->
                            spacing.Line <- StringValue "360"
                            spacing.LineRule <- EnumValue Wordprocessing.LineSpacingRuleValues.Auto
                        | DoubleSpacing ->
                            spacing.Line <- StringValue "480"
                            spacing.LineRule <- EnumValue Wordprocessing.LineSpacingRuleValues.Auto
                        | MultipleSpacing factor ->
                            spacing.Line <- StringValue(string (int (Math.Round(factor * 240.0))))
                            spacing.LineRule <- EnumValue Wordprocessing.LineSpacingRuleValues.Auto
                        | AtLeastSpacing pts ->
                            spacing.Line <- StringValue(string (twentiethsOfPoint pts))
                            spacing.LineRule <- EnumValue Wordprocessing.LineSpacingRuleValues.AtLeast
                        | ExactlySpacing pts ->
                            spacing.Line <- StringValue(string (twentiethsOfPoint pts))
                            spacing.LineRule <- EnumValue Wordprocessing.LineSpacingRuleValues.Exact)

                    pPr.SpacingBetweenLines <- spacing

                f.Indentation |> Option.iter (fun ind -> pPr.Indentation <- indentationToW ind)

                if f.KeepWithNext then
                    pPr.KeepNext <- Wordprocessing.KeepNext()

                if f.PageBreakBefore then
                    pPr.PageBreakBefore <- Wordprocessing.PageBreakBefore())

            Some pPr

    let paragraphFormatOfProperties (pPr: Wordprocessing.ParagraphProperties option) : ParagraphFormat option =
        match pPr with
        | None -> None
        | Some pPr ->
            let hasAny =
                isNull pPr.Justification |> not
                || isNull pPr.SpacingBetweenLines |> not
                || isNull pPr.Indentation |> not
                || isNull pPr.KeepNext |> not
                || isNull pPr.PageBreakBefore |> not

            if not hasAny then
                None
            else
                let lineSpacing =
                    if isNull pPr.SpacingBetweenLines then
                        None
                    else
                        let sp = pPr.SpacingBetweenLines

                        match sp.Line |> Option.ofObj |> Option.map (fun v -> v.Value) with
                        | None -> None
                        | Some lineStr ->
                            let lineVal = int lineStr
                            let rule = sp.LineRule |> Option.ofObj |> Option.map (fun v -> v.Value)

                            if rule = Some Wordprocessing.LineSpacingRuleValues.AtLeast then
                                Some(AtLeastSpacing(pointsOfTwentieths lineVal))
                            elif rule = Some Wordprocessing.LineSpacingRuleValues.Exact then
                                Some(ExactlySpacing(pointsOfTwentieths lineVal))
                            elif lineVal = 240 then
                                Some SingleSpacing
                            elif lineVal = 360 then
                                Some OnePointFiveSpacing
                            elif lineVal = 480 then
                                Some DoubleSpacing
                            else
                                Some(MultipleSpacing(float lineVal / 240.0))

                Some
                    { Alignment =
                        if isNull pPr.Justification then
                            None
                        else
                            pPr.Justification.Val
                            |> Option.ofObj
                            |> Option.map (fun v ->
                                if v.Value = Wordprocessing.JustificationValues.Center then AlignCenter
                                elif v.Value = Wordprocessing.JustificationValues.Right then AlignRight
                                elif v.Value = Wordprocessing.JustificationValues.Both then AlignJustify
                                else AlignLeft)
                      SpacingBefore =
                        if isNull pPr.SpacingBetweenLines then
                            None
                        else
                            pPr.SpacingBetweenLines.Before
                            |> Option.ofObj
                            |> Option.map (fun v -> pointsOfTwentieths (int v.Value))
                      SpacingAfter =
                        if isNull pPr.SpacingBetweenLines then
                            None
                        else
                            pPr.SpacingBetweenLines.After
                            |> Option.ofObj
                            |> Option.map (fun v -> pointsOfTwentieths (int v.Value))
                      LineSpacing = lineSpacing
                      Indentation = if isNull pPr.Indentation then None else Some(indentationOfW pPr.Indentation)
                      KeepWithNext = not (isNull pPr.KeepNext)
                      PageBreakBefore = not (isNull pPr.PageBreakBefore) }

    let styleIdOfParagraphProperties (pPr: Wordprocessing.ParagraphProperties option) : string option =
        match pPr with
        | Some pPr when not (isNull pPr.ParagraphStyleId) ->
            pPr.ParagraphStyleId.Val |> Option.ofObj |> Option.map (fun v -> v.Value)
        | _ -> None

    // --- Named styles (styles.xml) -------------------------------------------------------

    /// Ensures `"Normal"` is present (Word requires a default paragraph style to exist) and
    /// writes every `StyleDefinition` given, in order - this DSL doesn't resolve `BasedOn`
    /// chains itself (see `NamedStyles.StyleDefinition`'s own doc comment), it only ensures
    /// each referenced id gets a real element.
    let stylesToOpenXml (definitions: StyleDefinition list) : Wordprocessing.Styles =
        let defs =
            if definitions |> List.exists (fun d -> d.Id = "Normal") then
                definitions
            else
                BuiltInStyles.normal :: definitions

        let styleElements =
            defs
            |> List.map (fun d ->
                let s =
                    Wordprocessing.Style(
                        Type =
                            EnumValue(
                                match d.Type with
                                | ParagraphStyleType -> Wordprocessing.StyleValues.Paragraph
                                | CharacterStyleType -> Wordprocessing.StyleValues.Character
                            ),
                        StyleId = StringValue d.Id
                    )

                s.StyleName <- Wordprocessing.StyleName(Val = StringValue d.Name)
                d.BasedOn |> Option.iter (fun b -> s.BasedOn <- Wordprocessing.BasedOn(Val = StringValue b))

                if d.Id = "Normal" then
                    s.Default <- OnOffValue true

                d.ParaFormat
                |> Option.iter (fun pf ->
                    match paragraphPropertiesOf None (Some pf) with
                    | Some pPr -> s.StyleParagraphProperties <- Wordprocessing.StyleParagraphProperties(pPr.ChildElements |> Seq.map (fun c -> c.CloneNode true))
                    | None -> ())

                d.RunFormat
                |> Option.iter (fun rf ->
                    match runPropertiesOf (Some rf) None with
                    | Some rPr -> s.StyleRunProperties <- Wordprocessing.StyleRunProperties(rPr.ChildElements |> Seq.map (fun c -> c.CloneNode true))
                    | None -> ())

                s :> OpenXmlElement)

        Wordprocessing.Styles(styleElements)

    let stylesOfOpenXml (styles: Wordprocessing.Styles option) : StyleDefinition list =
        match styles with
        | None -> []
        | Some styles ->
            styles.Elements<Wordprocessing.Style>()
            |> Seq.map (fun s ->
                let paraFormat =
                    if isNull s.StyleParagraphProperties then
                        None
                    else
                        let pPr = Wordprocessing.ParagraphProperties(s.StyleParagraphProperties.ChildElements |> Seq.map (fun c -> c.CloneNode true))
                        paragraphFormatOfProperties (Some pPr)

                let runFormat =
                    if isNull s.StyleRunProperties then
                        None
                    else
                        let rPr = Wordprocessing.RunProperties(s.StyleRunProperties.ChildElements |> Seq.map (fun c -> c.CloneNode true))
                        runStyleOfProperties (Some rPr)

                { Id = s.StyleId.Value
                  Name = (if isNull s.StyleName then s.StyleId.Value else s.StyleName.Val.Value)
                  Type =
                    if not (isNull s.Type) && s.Type.Value = Wordprocessing.StyleValues.Character then
                        CharacterStyleType
                    else
                        ParagraphStyleType
                  BasedOn = (if isNull s.BasedOn then None else s.BasedOn.Val |> Option.ofObj |> Option.map (fun v -> v.Value))
                  RunFormat = runFormat
                  ParaFormat = paraFormat })
            |> List.ofSeq
