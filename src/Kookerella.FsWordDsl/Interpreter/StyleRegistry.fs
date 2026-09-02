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

    /// The plain hex/`"auto"` fallback value every consumer writes to `w:val`/`w:fill`
    /// regardless of `Color` case - a `Theme` color's own `Fallback` is used here, same as
    /// real Word writes a computed value alongside its own `w:themeColor`/`w:themeFill`.
    /// See `applyThemeToColor`/`applyThemeToShadingFill`/`colorOfRunColor`/
    /// `colorOfShadingFill` below for the theme-token-preserving read/write path (run color
    /// and shading/fill only - see `Color.Theme`'s own doc comment on borders).
    let colorToHex (c: Color) : string =
        match c with
        | Rgb(r, g, b) -> sprintf "%02X%02X%02X" r g b
        | Auto -> "auto"
        | Theme(_, (r, g, b), _, _) -> sprintf "%02X%02X%02X" r g b

    let colorOfHex (hex: string) : Color =
        if String.Equals(hex, "auto", StringComparison.OrdinalIgnoreCase) then
            Auto
        else
            let n = Convert.ToInt32(hex, 16)
            Rgb(byte ((n >>> 16) &&& 0xFF), byte ((n >>> 8) &&& 0xFF), byte (n &&& 0xFF))

    let private themeColorKindToW (k: ThemeColorKind) : Wordprocessing.ThemeColorValues =
        match k with
        | Dark1Theme -> Wordprocessing.ThemeColorValues.Dark1
        | Light1Theme -> Wordprocessing.ThemeColorValues.Light1
        | Dark2Theme -> Wordprocessing.ThemeColorValues.Dark2
        | Light2Theme -> Wordprocessing.ThemeColorValues.Light2
        | Accent1Theme -> Wordprocessing.ThemeColorValues.Accent1
        | Accent2Theme -> Wordprocessing.ThemeColorValues.Accent2
        | Accent3Theme -> Wordprocessing.ThemeColorValues.Accent3
        | Accent4Theme -> Wordprocessing.ThemeColorValues.Accent4
        | Accent5Theme -> Wordprocessing.ThemeColorValues.Accent5
        | Accent6Theme -> Wordprocessing.ThemeColorValues.Accent6
        | HyperlinkTheme -> Wordprocessing.ThemeColorValues.Hyperlink
        | FollowedHyperlinkTheme -> Wordprocessing.ThemeColorValues.FollowedHyperlink
        | Background1Theme -> Wordprocessing.ThemeColorValues.Background1
        | Text1Theme -> Wordprocessing.ThemeColorValues.Text1
        | Background2Theme -> Wordprocessing.ThemeColorValues.Background2
        | Text2Theme -> Wordprocessing.ThemeColorValues.Text2

    let private themeColorKindOfW (v: Wordprocessing.ThemeColorValues) : ThemeColorKind =
        if v = Wordprocessing.ThemeColorValues.Dark1 then Dark1Theme
        elif v = Wordprocessing.ThemeColorValues.Light1 then Light1Theme
        elif v = Wordprocessing.ThemeColorValues.Dark2 then Dark2Theme
        elif v = Wordprocessing.ThemeColorValues.Light2 then Light2Theme
        elif v = Wordprocessing.ThemeColorValues.Accent1 then Accent1Theme
        elif v = Wordprocessing.ThemeColorValues.Accent2 then Accent2Theme
        elif v = Wordprocessing.ThemeColorValues.Accent3 then Accent3Theme
        elif v = Wordprocessing.ThemeColorValues.Accent4 then Accent4Theme
        elif v = Wordprocessing.ThemeColorValues.Accent5 then Accent5Theme
        elif v = Wordprocessing.ThemeColorValues.Accent6 then Accent6Theme
        elif v = Wordprocessing.ThemeColorValues.Hyperlink then HyperlinkTheme
        elif v = Wordprocessing.ThemeColorValues.FollowedHyperlink then FollowedHyperlinkTheme
        elif v = Wordprocessing.ThemeColorValues.Background1 then Background1Theme
        elif v = Wordprocessing.ThemeColorValues.Text1 then Text1Theme
        elif v = Wordprocessing.ThemeColorValues.Background2 then Background2Theme
        else Text2Theme

    /// OOXML's own `w:themeTint`/`w:themeShade`/`w:themeFillTint`/`w:themeFillShade` are a
    /// single byte (`00`-`FF`) where `FF` = 100% - this DSL exposes that as `0.0`-`1.0`
    /// instead, matching the slider Word's own UI shows.
    let private tintToHexByte (pct: float) : string = sprintf "%02X" (byte (Math.Round(pct * 255.0)))
    let private hexByteToTint (hex: string) : float = float (Convert.ToInt32(hex, 16)) / 255.0

    /// Stamps `w:themeColor`/`w:themeTint`/`w:themeShade` onto an already-built `w:color`
    /// element when `c` is `Theme` - a no-op for `Rgb`/`Auto` (`colorToHex` above already
    /// wrote the plain `w:val` either way).
    let applyThemeToColor (rc: Wordprocessing.Color) (c: Color) : unit =
        match c with
        | Theme(kind, _, tint, shade) ->
            rc.ThemeColor <- EnumValue(themeColorKindToW kind)
            tint |> Option.iter (fun t -> rc.ThemeTint <- StringValue(tintToHexByte t))
            shade |> Option.iter (fun s -> rc.ThemeShade <- StringValue(tintToHexByte s))
        | _ -> ()

    /// The `w:shd`-side equivalent of `applyThemeToColor`, for the fill/background theme
    /// attributes (`w:themeFill`/`w:themeFillTint`/`w:themeFillShade`) - `w:shd`'s
    /// foreground `w:themeColor` attributes are unused here since this DSL always writes a
    /// plain `w:color="auto"` foreground for shading (see `paragraphPropertiesOf` etc.).
    let applyThemeToShadingFill (sh: Wordprocessing.Shading) (c: Color) : unit =
        match c with
        | Theme(kind, _, tint, shade) ->
            sh.ThemeFill <- EnumValue(themeColorKindToW kind)
            tint |> Option.iter (fun t -> sh.ThemeFillTint <- StringValue(tintToHexByte t))
            shade |> Option.iter (fun s -> sh.ThemeFillShade <- StringValue(tintToHexByte s))
        | _ -> ()

    /// The inverse of `applyThemeToColor` plus `colorToHex`/`colorOfHex` together - reads a
    /// `w:color` element back as `Theme` when it carries a theme token, `Rgb`/`Auto`
    /// (via `colorOfHex`) otherwise.
    let colorOfRunColor (rc: Wordprocessing.Color) : Color =
        let hex = if isNull rc.Val then "auto" else rc.Val.Value

        match rc.ThemeColor |> Option.ofObj with
        | Some tc ->
            Theme(
                themeColorKindOfW tc.Value,
                (match colorOfHex hex with
                 | Rgb(r, g, b) -> r, g, b
                 | _ -> 0uy, 0uy, 0uy),
                rc.ThemeTint |> Option.ofObj |> Option.map (fun v -> hexByteToTint v.Value),
                rc.ThemeShade |> Option.ofObj |> Option.map (fun v -> hexByteToTint v.Value)
            )
        | None -> colorOfHex hex

    /// The `w:shd`-side equivalent of `colorOfRunColor`, reading `w:fill`/`w:themeFill`/
    /// `w:themeFillTint`/`w:themeFillShade` back.
    let colorOfShadingFill (sh: Wordprocessing.Shading) : Color =
        let hex = if isNull sh.Fill then "auto" else sh.Fill.Value

        match sh.ThemeFill |> Option.ofObj with
        | Some tf ->
            Theme(
                themeColorKindOfW tf.Value,
                (match colorOfHex hex with
                 | Rgb(r, g, b) -> r, g, b
                 | _ -> 0uy, 0uy, 0uy),
                sh.ThemeFillTint |> Option.ofObj |> Option.map (fun v -> hexByteToTint v.Value),
                sh.ThemeFillShade |> Option.ofObj |> Option.map (fun v -> hexByteToTint v.Value)
            )
        | None -> colorOfHex hex

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

    /// Shared by table borders (`Writer.fs`'s `TableBorders`/`TableCellBorders`) and
    /// paragraph borders (`w:pBdr`, below) - both contexts reuse the very same
    /// `TopBorder`/`LeftBorder`/`BottomBorder`/`RightBorder`/`InsideHorizontalBorder`/
    /// `InsideVerticalBorder` SDK classes (confirmed by reflection, not assumed - see this
    /// repo's `CLAUDE.md` on why that verification step matters for this SDK).
    let borderSideToTop (side: BorderSide) : Wordprocessing.TopBorder =
        let b = Wordprocessing.TopBorder(Val = EnumValue(borderLineStyleToW side.Style), Size = UInt32Value(borderSideWidthEighths side), Space = UInt32Value 0u)
        side.Color |> Option.iter (fun c -> b.Color <- StringValue(colorToHex c))
        b

    let borderSideToBottom (side: BorderSide) : Wordprocessing.BottomBorder =
        let b = Wordprocessing.BottomBorder(Val = EnumValue(borderLineStyleToW side.Style), Size = UInt32Value(borderSideWidthEighths side), Space = UInt32Value 0u)
        side.Color |> Option.iter (fun c -> b.Color <- StringValue(colorToHex c))
        b

    let borderSideToLeft (side: BorderSide) : Wordprocessing.LeftBorder =
        let b = Wordprocessing.LeftBorder(Val = EnumValue(borderLineStyleToW side.Style), Size = UInt32Value(borderSideWidthEighths side), Space = UInt32Value 0u)
        side.Color |> Option.iter (fun c -> b.Color <- StringValue(colorToHex c))
        b

    let borderSideToRight (side: BorderSide) : Wordprocessing.RightBorder =
        let b = Wordprocessing.RightBorder(Val = EnumValue(borderLineStyleToW side.Style), Size = UInt32Value(borderSideWidthEighths side), Space = UInt32Value 0u)
        side.Color |> Option.iter (fun c -> b.Color <- StringValue(colorToHex c))
        b

    let borderSideToInsideH (side: BorderSide) : Wordprocessing.InsideHorizontalBorder =
        let b = Wordprocessing.InsideHorizontalBorder(Val = EnumValue(borderLineStyleToW side.Style), Size = UInt32Value(borderSideWidthEighths side), Space = UInt32Value 0u)
        side.Color |> Option.iter (fun c -> b.Color <- StringValue(colorToHex c))
        b

    let borderSideToInsideV (side: BorderSide) : Wordprocessing.InsideVerticalBorder =
        let b = Wordprocessing.InsideVerticalBorder(Val = EnumValue(borderLineStyleToW side.Style), Size = UInt32Value(borderSideWidthEighths side), Space = UInt32Value 0u)
        side.Color |> Option.iter (fun c -> b.Color <- StringValue(colorToHex c))
        b

    let borderSideOfTop (b: Wordprocessing.TopBorder) : BorderSide =
        { Style = borderLineStyleOfW b.Val.Value
          Width = b.Size |> Option.ofObj |> Option.map (fun v -> borderSideOfWidthEighths v.Value)
          Color = b.Color |> Option.ofObj |> Option.map (fun v -> colorOfHex v.Value) }

    let borderSideOfBottom (b: Wordprocessing.BottomBorder) : BorderSide =
        { Style = borderLineStyleOfW b.Val.Value
          Width = b.Size |> Option.ofObj |> Option.map (fun v -> borderSideOfWidthEighths v.Value)
          Color = b.Color |> Option.ofObj |> Option.map (fun v -> colorOfHex v.Value) }

    let borderSideOfLeft (b: Wordprocessing.LeftBorder) : BorderSide =
        { Style = borderLineStyleOfW b.Val.Value
          Width = b.Size |> Option.ofObj |> Option.map (fun v -> borderSideOfWidthEighths v.Value)
          Color = b.Color |> Option.ofObj |> Option.map (fun v -> colorOfHex v.Value) }

    let borderSideOfRight (b: Wordprocessing.RightBorder) : BorderSide =
        { Style = borderLineStyleOfW b.Val.Value
          Width = b.Size |> Option.ofObj |> Option.map (fun v -> borderSideOfWidthEighths v.Value)
          Color = b.Color |> Option.ofObj |> Option.map (fun v -> colorOfHex v.Value) }

    let borderSideOfInsideH (b: Wordprocessing.InsideHorizontalBorder) : BorderSide =
        { Style = borderLineStyleOfW b.Val.Value
          Width = b.Size |> Option.ofObj |> Option.map (fun v -> borderSideOfWidthEighths v.Value)
          Color = b.Color |> Option.ofObj |> Option.map (fun v -> colorOfHex v.Value) }

    let borderSideOfInsideV (b: Wordprocessing.InsideVerticalBorder) : BorderSide =
        { Style = borderLineStyleOfW b.Val.Value
          Width = b.Size |> Option.ofObj |> Option.map (fun v -> borderSideOfWidthEighths v.Value)
          Color = b.Color |> Option.ofObj |> Option.map (fun v -> colorOfHex v.Value) }

    let paragraphBordersToW (b: BorderStyle) : Wordprocessing.ParagraphBorders =
        let pBdr = Wordprocessing.ParagraphBorders()
        b.Top |> Option.iter (fun s -> pBdr.TopBorder <- borderSideToTop s)
        b.Bottom |> Option.iter (fun s -> pBdr.BottomBorder <- borderSideToBottom s)
        b.Left |> Option.iter (fun s -> pBdr.LeftBorder <- borderSideToLeft s)
        b.Right |> Option.iter (fun s -> pBdr.RightBorder <- borderSideToRight s)
        pBdr

    let paragraphBordersOfW (pBdr: Wordprocessing.ParagraphBorders) : BorderStyle =
        { Top = pBdr.TopBorder |> Option.ofObj |> Option.map borderSideOfTop
          Bottom = pBdr.BottomBorder |> Option.ofObj |> Option.map borderSideOfBottom
          Left = pBdr.LeftBorder |> Option.ofObj |> Option.map borderSideOfLeft
          Right = pBdr.RightBorder |> Option.ofObj |> Option.map borderSideOfRight }

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
                |> Option.iter (fun c ->
                    let rc = Wordprocessing.Color(Val = StringValue(colorToHex c))
                    applyThemeToColor rc c
                    rPr.Color <- rc)

                s.Highlight
                |> Option.iter (fun h -> rPr.Highlight <- Wordprocessing.Highlight(Val = EnumValue(highlightToW h)))

                s.VerticalPosition
                |> Option.iter (fun vp ->
                    let v =
                        match vp with
                        | Superscript -> Wordprocessing.VerticalPositionValues.Superscript
                        | Subscript -> Wordprocessing.VerticalPositionValues.Subscript

                    rPr.VerticalTextAlignment <- Wordprocessing.VerticalTextAlignment(Val = EnumValue v))

                if s.SmallCaps then
                    rPr.SmallCaps <- Wordprocessing.SmallCaps()

                if s.AllCaps then
                    rPr.Caps <- Wordprocessing.Caps()

                if s.Hidden then
                    rPr.Vanish <- Wordprocessing.Vanish())

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
                || isNull rPr.SmallCaps |> not
                || isNull rPr.Caps |> not
                || isNull rPr.Vanish |> not

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
                        Color = if isNull rPr.Color then None else Some(colorOfRunColor rPr.Color)
                        Highlight =
                            if isNull rPr.Highlight then None
                            else rPr.Highlight.Val |> Option.ofObj |> Option.bind (fun v -> highlightOfW v.Value)
                        VerticalPosition =
                            if isNull rPr.VerticalTextAlignment then None
                            else
                                rPr.VerticalTextAlignment.Val
                                |> Option.ofObj
                                |> Option.map (fun v ->
                                    if v.Value = Wordprocessing.VerticalPositionValues.Superscript then Superscript else Subscript)
                        SmallCaps = not (isNull rPr.SmallCaps)
                        AllCaps = not (isNull rPr.Caps)
                        Hidden = not (isNull rPr.Vanish) }

    let styleIdOfProperties (rPr: Wordprocessing.RunProperties option) : string option =
        match rPr with
        | Some rPr when not (isNull rPr.RunStyle) -> rPr.RunStyle.Val |> Option.ofObj |> Option.map (fun v -> v.Value)
        | _ -> None

    let private twentiethsOfPoint (pts: float) : int = int (Math.Round(pts * 20.0))
    let private pointsOfTwentieths (v: int) : float = float v / 20.0

    // --- Tab stops -----------------------------------------------------------------------

    let tabStopAlignmentToW (a: TabStopAlignment) : Wordprocessing.TabStopValues =
        match a with
        | LeftTab -> Wordprocessing.TabStopValues.Left
        | CenterTab -> Wordprocessing.TabStopValues.Center
        | RightTab -> Wordprocessing.TabStopValues.Right
        | DecimalTab -> Wordprocessing.TabStopValues.Decimal
        | BarTab -> Wordprocessing.TabStopValues.Bar
        | OtherTabAlignment raw -> Wordprocessing.TabStopValues raw

    let tabStopAlignmentOfW (v: Wordprocessing.TabStopValues) : TabStopAlignment =
        if v = Wordprocessing.TabStopValues.Left then LeftTab
        elif v = Wordprocessing.TabStopValues.Center then CenterTab
        elif v = Wordprocessing.TabStopValues.Right then RightTab
        elif v = Wordprocessing.TabStopValues.Decimal then DecimalTab
        elif v = Wordprocessing.TabStopValues.Bar then BarTab
        else OtherTabAlignment(v.ToString())

    let tabLeaderToW (l: TabLeader) : Wordprocessing.TabStopLeaderCharValues =
        match l with
        | NoLeader -> Wordprocessing.TabStopLeaderCharValues.None
        | DotLeader -> Wordprocessing.TabStopLeaderCharValues.Dot
        | HyphenLeader -> Wordprocessing.TabStopLeaderCharValues.Hyphen
        | UnderscoreLeader -> Wordprocessing.TabStopLeaderCharValues.Underscore
        | HeavyLeader -> Wordprocessing.TabStopLeaderCharValues.Heavy
        | MiddleDotLeader -> Wordprocessing.TabStopLeaderCharValues.MiddleDot

    let tabLeaderOfW (v: Wordprocessing.TabStopLeaderCharValues) : TabLeader =
        if v = Wordprocessing.TabStopLeaderCharValues.Dot then DotLeader
        elif v = Wordprocessing.TabStopLeaderCharValues.Hyphen then HyphenLeader
        elif v = Wordprocessing.TabStopLeaderCharValues.Underscore then UnderscoreLeader
        elif v = Wordprocessing.TabStopLeaderCharValues.Heavy then HeavyLeader
        elif v = Wordprocessing.TabStopLeaderCharValues.MiddleDot then MiddleDotLeader
        else NoLeader

    let tabStopToW (t: TabStop) : Wordprocessing.TabStop =
        Wordprocessing.TabStop(Val = EnumValue(tabStopAlignmentToW t.Alignment), Position = Int32Value(twentiethsOfPoint t.Position), Leader = EnumValue(tabLeaderToW t.Leader))

    let tabStopOfW (t: Wordprocessing.TabStop) : TabStop =
        { Position = t.Position |> Option.ofObj |> Option.map (fun v -> pointsOfTwentieths v.Value) |> Option.defaultValue 0.0
          Alignment = t.Val |> Option.ofObj |> Option.map (fun v -> tabStopAlignmentOfW v.Value) |> Option.defaultValue LeftTab
          Leader = t.Leader |> Option.ofObj |> Option.map (fun v -> tabLeaderOfW v.Value) |> Option.defaultValue NoLeader }

    let tabsToW (stops: TabStop list) : Wordprocessing.Tabs =
        let tabs = Wordprocessing.Tabs()
        stops |> List.iter (fun t -> tabs.AppendChild(tabStopToW t) |> ignore)
        tabs

    let tabsOfW (tabs: Wordprocessing.Tabs) : TabStop list =
        tabs.Elements<Wordprocessing.TabStop>() |> Seq.map tabStopOfW |> List.ofSeq

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
                    pPr.PageBreakBefore <- Wordprocessing.PageBreakBefore()

                f.Borders |> Option.iter (fun b -> pPr.ParagraphBorders <- paragraphBordersToW b)

                f.Shading
                |> Option.iter (fun c ->
                    let sh = Wordprocessing.Shading(Val = EnumValue Wordprocessing.ShadingPatternValues.Clear, Color = StringValue "auto", Fill = StringValue(colorToHex c))
                    applyThemeToShadingFill sh c
                    pPr.Shading <- sh)

                if not f.TabStops.IsEmpty then
                    pPr.Tabs <- tabsToW f.TabStops)

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
                || isNull pPr.ParagraphBorders |> not
                || isNull pPr.Shading |> not
                || isNull pPr.Tabs |> not

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
                      PageBreakBefore = not (isNull pPr.PageBreakBefore)
                      Borders = if isNull pPr.ParagraphBorders then None else Some(paragraphBordersOfW pPr.ParagraphBorders)
                      Shading = if isNull pPr.Shading then None else Some(colorOfShadingFill pPr.Shading)
                      TabStops = if isNull pPr.Tabs then [] else tabsOfW pPr.Tabs }

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

    /// Only paragraph/character styles - `w:type="table"` entries are `Document.TableStyles`
    /// instead (see `tableStylesOfOpenXml` below), and would otherwise get misread here as
    /// spurious `ParagraphStyleType` entries (this DSL doesn't write `w:type="numbering"`
    /// styles itself, but skips those too rather than misreading them the same way).
    let stylesOfOpenXml (styles: Wordprocessing.Styles option) : StyleDefinition list =
        match styles with
        | None -> []
        | Some styles ->
            styles.Elements<Wordprocessing.Style>()
            |> Seq.filter (fun s -> isNull s.Type || s.Type.Value = Wordprocessing.StyleValues.Paragraph || s.Type.Value = Wordprocessing.StyleValues.Character)
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

    // --- Table style definitions (styles.xml, w:type="table") ----------------------------

    let private tableBordersToOpenXml (b: TableBorders) : Wordprocessing.TableBorders =
        let tb = Wordprocessing.TableBorders()
        b.Outer.Top |> Option.iter (fun s -> tb.TopBorder <- borderSideToTop s)
        b.Outer.Bottom |> Option.iter (fun s -> tb.BottomBorder <- borderSideToBottom s)
        b.Outer.Left |> Option.iter (fun s -> tb.LeftBorder <- borderSideToLeft s)
        b.Outer.Right |> Option.iter (fun s -> tb.RightBorder <- borderSideToRight s)
        b.InsideHorizontal |> Option.iter (fun s -> tb.InsideHorizontalBorder <- borderSideToInsideH s)
        b.InsideVertical |> Option.iter (fun s -> tb.InsideVerticalBorder <- borderSideToInsideV s)
        tb

    let private tableBordersOfOpenXml (tb: Wordprocessing.TableBorders) : TableBorders =
        { Outer =
            { Left = tb.LeftBorder |> Option.ofObj |> Option.map borderSideOfLeft
              Right = tb.RightBorder |> Option.ofObj |> Option.map borderSideOfRight
              Top = tb.TopBorder |> Option.ofObj |> Option.map borderSideOfTop
              Bottom = tb.BottomBorder |> Option.ofObj |> Option.map borderSideOfBottom }
          InsideHorizontal = tb.InsideHorizontalBorder |> Option.ofObj |> Option.map borderSideOfInsideH
          InsideVertical = tb.InsideVerticalBorder |> Option.ofObj |> Option.map borderSideOfInsideV }

    let private tableStyleRegionToOpenXml (kind: Wordprocessing.TableStyleOverrideValues) (region: TableStyleRegion) : Wordprocessing.TableStyleProperties option =
        if region = TableStyleRegion.None then
            None
        else
            let tsp = Wordprocessing.TableStyleProperties(Type = EnumValue kind)

            region.RunFormat
            |> Option.iter (fun rf ->
                match runPropertiesOf (Some rf) None with
                | Some rPr -> tsp.RunPropertiesBaseStyle <- Wordprocessing.RunPropertiesBaseStyle(rPr.ChildElements |> Seq.map (fun c -> c.CloneNode true))
                | None -> ())

            region.ParaFormat
            |> Option.iter (fun pf ->
                match paragraphPropertiesOf None (Some pf) with
                | Some pPr -> tsp.StyleParagraphProperties <- Wordprocessing.StyleParagraphProperties(pPr.ChildElements |> Seq.map (fun c -> c.CloneNode true))
                | None -> ())

            region.CellShading
            |> Option.iter (fun c ->
                let tcPr = Wordprocessing.TableStyleConditionalFormattingTableCellProperties()
                let sh = Wordprocessing.Shading(Val = EnumValue Wordprocessing.ShadingPatternValues.Clear, Color = StringValue "auto", Fill = StringValue(colorToHex c))
                applyThemeToShadingFill sh c
                tcPr.Shading <- sh
                tsp.TableStyleConditionalFormattingTableCellProperties <- tcPr)

            Some tsp

    let private tableStyleRegionOfOpenXml (tsp: Wordprocessing.TableStyleProperties) : TableStyleRegion =
        let runFormat =
            tsp.RunPropertiesBaseStyle
            |> Option.ofObj
            |> Option.bind (fun rpb -> runStyleOfProperties (Some(Wordprocessing.RunProperties(rpb.ChildElements |> Seq.map (fun c -> c.CloneNode true)))))

        let paraFormat =
            tsp.StyleParagraphProperties
            |> Option.ofObj
            |> Option.bind (fun spp -> paragraphFormatOfProperties (Some(Wordprocessing.ParagraphProperties(spp.ChildElements |> Seq.map (fun c -> c.CloneNode true)))))

        let cellShading =
            tsp.TableStyleConditionalFormattingTableCellProperties
            |> Option.ofObj
            |> Option.bind (fun tcp -> tcp.Shading |> Option.ofObj)
            |> Option.map colorOfShadingFill

        { RunFormat = runFormat; ParaFormat = paraFormat; CellShading = cellShading }

    /// Builds one `w:style[@w:type='table']` element per definition, to append to the same
    /// `Wordprocessing.Styles` collection `stylesToOpenXml` builds - kept as a separate
    /// function (rather than folded into `stylesToOpenXml`) since `Document.TableStyles` is
    /// its own field, not part of `Document.Styles`.
    let tableStylesToOpenXml (definitions: TableStyleDefinition list) : Wordprocessing.Style list =
        definitions
        |> List.map (fun d ->
            let s = Wordprocessing.Style(Type = EnumValue Wordprocessing.StyleValues.Table, StyleId = StringValue d.Id)
            s.StyleName <- Wordprocessing.StyleName(Val = StringValue d.Name)
            d.BasedOn |> Option.iter (fun b -> s.BasedOn <- Wordprocessing.BasedOn(Val = StringValue b))

            d.Borders
            |> Option.iter (fun b ->
                let stp = Wordprocessing.StyleTableProperties()
                stp.TableBorders <- tableBordersToOpenXml b
                s.StyleTableProperties <- stp)

            [ Wordprocessing.TableStyleOverrideValues.WholeTable, d.WholeTable
              Wordprocessing.TableStyleOverrideValues.FirstRow, d.FirstRow
              Wordprocessing.TableStyleOverrideValues.LastRow, d.LastRow
              Wordprocessing.TableStyleOverrideValues.FirstColumn, d.FirstColumn
              Wordprocessing.TableStyleOverrideValues.LastColumn, d.LastColumn
              Wordprocessing.TableStyleOverrideValues.Band1Horizontal, d.BandedRow
              Wordprocessing.TableStyleOverrideValues.Band1Vertical, d.BandedColumn
              Wordprocessing.TableStyleOverrideValues.NorthEastCell, d.NorthEastCell
              Wordprocessing.TableStyleOverrideValues.NorthWestCell, d.NorthWestCell
              Wordprocessing.TableStyleOverrideValues.SouthEastCell, d.SouthEastCell
              Wordprocessing.TableStyleOverrideValues.SouthWestCell, d.SouthWestCell ]
            |> List.iter (fun (kind, region) -> tableStyleRegionToOpenXml kind region |> Option.iter (fun tsp -> s.AppendChild(tsp) |> ignore))

            s)

    /// The inverse of `tableStylesToOpenXml` - only `w:type="table"` entries, see
    /// `stylesOfOpenXml`'s own note on why the two functions each filter to their own kind.
    let tableStylesOfOpenXml (styles: Wordprocessing.Styles option) : TableStyleDefinition list =
        match styles with
        | None -> []
        | Some styles ->
            styles.Elements<Wordprocessing.Style>()
            |> Seq.filter (fun s -> not (isNull s.Type) && s.Type.Value = Wordprocessing.StyleValues.Table)
            |> Seq.map (fun s ->
                let regionOf (kind: Wordprocessing.TableStyleOverrideValues) =
                    s.Elements<Wordprocessing.TableStyleProperties>()
                    |> Seq.tryFind (fun t -> not (isNull t.Type) && t.Type.Value = kind)
                    |> Option.map tableStyleRegionOfOpenXml
                    |> Option.defaultValue TableStyleRegion.None

                { Id = s.StyleId.Value
                  Name = (if isNull s.StyleName then s.StyleId.Value else s.StyleName.Val.Value)
                  BasedOn = (if isNull s.BasedOn then None else s.BasedOn.Val |> Option.ofObj |> Option.map (fun v -> v.Value))
                  Borders =
                    s.StyleTableProperties
                    |> Option.ofObj
                    |> Option.bind (fun stp -> stp.TableBorders |> Option.ofObj)
                    |> Option.map tableBordersOfOpenXml
                  WholeTable = regionOf Wordprocessing.TableStyleOverrideValues.WholeTable
                  FirstRow = regionOf Wordprocessing.TableStyleOverrideValues.FirstRow
                  LastRow = regionOf Wordprocessing.TableStyleOverrideValues.LastRow
                  FirstColumn = regionOf Wordprocessing.TableStyleOverrideValues.FirstColumn
                  LastColumn = regionOf Wordprocessing.TableStyleOverrideValues.LastColumn
                  BandedRow = regionOf Wordprocessing.TableStyleOverrideValues.Band1Horizontal
                  BandedColumn = regionOf Wordprocessing.TableStyleOverrideValues.Band1Vertical
                  NorthEastCell = regionOf Wordprocessing.TableStyleOverrideValues.NorthEastCell
                  NorthWestCell = regionOf Wordprocessing.TableStyleOverrideValues.NorthWestCell
                  SouthEastCell = regionOf Wordprocessing.TableStyleOverrideValues.SouthEastCell
                  SouthWestCell = regionOf Wordprocessing.TableStyleOverrideValues.SouthWestCell })
            |> List.ofSeq
