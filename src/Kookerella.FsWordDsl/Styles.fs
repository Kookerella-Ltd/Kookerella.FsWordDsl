namespace Kookerella.FsWordDsl

/// Character (run) and paragraph formatting. Unlike Excel's `CellStyle` - which the
/// interpreter interns into a shared, indexed stylesheet - WordprocessingML writes direct
/// run/paragraph formatting (`w:rPr`/`w:pPr`) inline on each element, not through an index.
/// So there is no dedup/interning concept mirrored here; `Interpreter/StyleRegistry.fs`'s
/// job is narrower (see its own doc comment).
[<AutoOpen>]
module Styles =

    /// Word's twelve standard theme color slots (`w:themeColor`/`w:themeFill`) - resolved
    /// against whatever theme is attached to the document when opened in real Word. This
    /// DSL doesn't model theme parts (`word/theme/theme1.xml`) themselves - see `Color.
    /// Theme`'s own doc comment for how that's handled.
    type ThemeColorKind =
        | Dark1Theme
        | Light1Theme
        | Dark2Theme
        | Light2Theme
        | Accent1Theme
        | Accent2Theme
        | Accent3Theme
        | Accent4Theme
        | Accent5Theme
        | Accent6Theme
        | HyperlinkTheme
        | FollowedHyperlinkTheme
        | Background1Theme
        | Text1Theme
        | Background2Theme
        | Text2Theme

    /// A run/shading color. `Auto` matches OOXML's own "let the reader/theme decide" value
    /// (the default for new text in Word).
    type Color =
        | Rgb of red: byte * green: byte * blue: byte
        | Auto
        /// A theme-relative color (`w:themeColor` plus an explicit `Fallback` RGB - the same
        /// "always write a computed value alongside the theme token" convention real Word
        /// itself follows, since this DSL has no theme part to resolve `kind` against; a
        /// reader with no theme, or one this DSL doesn't understand, still sees `Fallback`).
        /// `Tint`/`Shade` lighten/darken it (0.0-1.0, matching Word's own tint/shade
        /// slider) - stored on the wire as a single byte, so an arbitrary `float` here
        /// round-trips to the nearest `/255`, not bit-for-bit (e.g. `0.5` reads back as
        /// `~0.502`); a value that's already an exact multiple of `1/255` (`0.2`, `0.4`,
        /// ...) round-trips exactly, same "the wire format is the one source of truth"
        /// posture `BorderSide.Width`'s own eighths-of-a-point rounding takes. Only
        /// modeled for run color and shading/fill backgrounds (`RunStyle.
        /// Color`, `ParagraphFormat.Shading`, `TableCellProps.Shading`, `TableStyleRegion.
        /// CellShading`) - a border's own color (`BorderSide.Color`) round-trips a `Theme`
        /// value as its `Fallback` RGB only, the theme token itself isn't preserved there
        /// (see MAPPING.md).
        | Theme of kind: ThemeColorKind * fallback: (byte * byte * byte) * tint: float option * shade: float option

    module Color =
        let black = Rgb(0uy, 0uy, 0uy)
        let white = Rgb(255uy, 255uy, 255uy)
        let red = Rgb(255uy, 0uy, 0uy)
        let green = Rgb(0uy, 128uy, 0uy)
        let blue = Rgb(0uy, 0uy, 255uy)
        let yellow = Rgb(255uy, 255uy, 0uy)

    /// WordprocessingML's `w:highlight` element only accepts this fixed, enumerated palette
    /// (Word's own "text highlight color" swatch) - unlike `Color`, arbitrary RGB is not
    /// valid here, so this is deliberately its own closed type rather than reusing `Color`.
    type HighlightColor =
        | HlYellow
        | HlGreen
        | HlCyan
        | HlMagenta
        | HlBlue
        | HlRed
        | HlDarkBlue
        | HlDarkCyan
        | HlDarkGreen
        | HlDarkMagenta
        | HlDarkRed
        | HlDarkYellow
        | HlDarkGray
        | HlLightGray
        | HlBlack

    type UnderlineStyle =
        | SingleUnderline
        | DoubleUnderline
        | ThickUnderline
        | DottedUnderline
        | DashedUnderline
        | WavyUnderline
        /// Preserves any other raw OOXML `w:u/@w:val` so reading and re-writing an existing
        /// document round-trips even for underline kinds this DSL doesn't author itself -
        /// same escape-hatch convention as Excel's `BorderLineStyle.Other`.
        | OtherUnderline of string

    type VerticalPosition =
        | Superscript
        | Subscript

    /// Direct/inline character formatting - written straight onto the run's own `w:rPr`,
    /// never interned or deduplicated (see this module's own doc comment).
    type RunStyle =
        { FontFamily: string option
          /// Points.
          Size: float option
          Bold: bool
          Italic: bool
          Underline: UnderlineStyle option
          Strikethrough: bool
          Color: Color option
          Highlight: HighlightColor option
          VerticalPosition: VerticalPosition option
          /// Renders lowercase letters as smaller uppercase ones (`w:smallCaps`) - distinct
          /// from `AllCaps`, which renders them as full-size uppercase instead.
          SmallCaps: bool
          /// Renders every letter as uppercase for display, without changing the run's own
          /// stored text (`w:caps`) - mutually exclusive with `SmallCaps` in real Word (only
          /// one visibly wins), but this DSL doesn't prevent setting both, same "trust the
          /// caller" posture the rest of this module takes.
          AllCaps: bool
          /// Text present in the document but not displayed or printed until unhidden
          /// (`w:vanish`) - distinct from `DocumentProtection`, which restricts editing
          /// rather than visibility.
          Hidden: bool }

        static member Default =
            { FontFamily = None
              Size = None
              Bold = false
              Italic = false
              Underline = None
              Strikethrough = false
              Color = None
              Highlight = None
              VerticalPosition = None
              SmallCaps = false
              AllCaps = false
              Hidden = false }

    type ParagraphAlignment =
        | AlignLeft
        | AlignCenter
        | AlignRight
        | AlignJustify

    /// All fields are in points. `FirstLine`/`Hanging` are mutually meaningful alternatives
    /// (Word's own `w:ind` allows both `firstLine` and `hanging` attributes, but setting
    /// both is contradictory in practice) - this DSL doesn't prevent setting both, it just
    /// writes whichever are present, the same "trust the caller" posture Excel takes on its
    /// own optional style fields.
    type Indentation =
        { Left: float option
          Right: float option
          FirstLine: float option
          Hanging: float option }

        static member None =
            { Left = None
              Right = None
              FirstLine = None
              Hanging = None }

    type LineSpacingRule =
        | SingleSpacing
        | OnePointFiveSpacing
        | DoubleSpacing
        /// Points - line height is at least this tall, growing to fit taller content.
        | AtLeastSpacing of points: float
        /// Points - line height is fixed exactly, regardless of content.
        | ExactlySpacing of points: float
        /// A multiple of single line spacing, e.g. `Multiple 1.15`.
        | MultipleSpacing of factor: float

    /// WordprocessingML defines many more border styles than this names explicitly (dashed
    /// variants, triple lines, 3-D effects, art borders, ...); `OtherLine` preserves the raw
    /// OOXML style name so reading and re-writing an existing document round-trips even for
    /// a style this DSL doesn't author itself - same convention as Excel's
    /// `BorderLineStyle.Other`. Deliberately not named `Single`/`Thick`/`Double` bare (which
    /// would collide with `LineSpacingRule`'s cases under `[<AutoOpen>]`).
    type BorderLineStyle =
        | SingleLine
        | ThickLine
        | DoubleLine
        | DottedLine
        | DashedLine
        | WaveLine
        | OtherLine of string

    /// `Width` is in points (OOXML's own border `sz` attribute is in eighths of a point;
    /// `Interpreter/Writer.fs` does that conversion) - `None` uses Word's own default weight.
    type BorderSide =
        { Style: BorderLineStyle
          Width: float option
          Color: Color option }

    /// Reused for both paragraph borders (`w:pBdr`) and table/cell borders (`w:tblBorders`/
    /// `w:tcBorders`) - same shape as Excel's `BorderStyle`, which is reused across cell and
    /// conditional-formatting borders the same way. `w:pBdr` also allows `between`/`bar`
    /// sides (a line between consecutive same-bordered paragraphs, and a vertical bar) -
    /// not modeled, same "narrow scope, document the gap" posture as `Tables.TableBorders`
    /// not modeling diagonal cell borders.
    type BorderStyle =
        { Left: BorderSide option
          Right: BorderSide option
          Top: BorderSide option
          Bottom: BorderSide option }

        static member None =
            { Left = None
              Right = None
              Top = None
              Bottom = None }

    /// How text lines up against a custom tab stop (`w:tab/@w:val`). `OtherTabAlignment`
    /// preserves any other raw OOXML value (e.g. `"start"`/`"end"`, the bidi-aware aliases
    /// for `Left`/`Right`, or `"num"`) so reading and re-writing an existing document
    /// round-trips even for a kind this DSL doesn't author itself - same escape-hatch
    /// convention as `BorderLineStyle.Other`.
    type TabStopAlignment =
        | LeftTab
        | CenterTab
        | RightTab
        | DecimalTab
        | BarTab
        | OtherTabAlignment of string

    /// The dotted/dashed/etc. fill Word draws between the previous text and a tab stop
    /// (`w:tab/@w:leader`) - most commonly seen leading up to a right-aligned page number in
    /// a table of contents.
    type TabLeader =
        | NoLeader
        | DotLeader
        | HyphenLeader
        | UnderscoreLeader
        | HeavyLeader
        | MiddleDotLeader

    /// One custom tab stop (`w:tab`). `Position` is in points, measured from the left
    /// margin (matching every other position field in this module).
    type TabStop =
        { Position: float
          Alignment: TabStopAlignment
          Leader: TabLeader }

    /// Direct/inline paragraph formatting - written straight onto the paragraph's own
    /// `w:pPr`. `StyleId` (on `Paragraph` itself, not here) supplies the named-style layer;
    /// this record is the override/direct-formatting layer on top of it, same relationship
    /// direct cell formatting has to (the non-existent, for Excel) named cell styles.
    type ParagraphFormat =
        { Alignment: ParagraphAlignment option
          SpacingBefore: float option
          SpacingAfter: float option
          LineSpacing: LineSpacingRule option
          Indentation: Indentation option
          KeepWithNext: bool
          PageBreakBefore: bool
          /// The paragraph's own border box (`w:pBdr`) - independent of any table border
          /// the paragraph might also sit inside.
          Borders: BorderStyle option
          /// Background fill behind the paragraph's text (`w:shd`) - same `Color` type
          /// `Tables.TableCellProps.Shading` uses for a table cell's own background.
          Shading: Color option
          /// Custom tab stops (`w:tabs`) - an empty list means "no custom tabs", not
          /// "clear Word's own default tab stops every half-inch" (this DSL doesn't author
          /// `w:val="clear"` entries, which exist only to override an inherited style's
          /// tab stops).
          TabStops: TabStop list }

        static member Default =
            { Alignment = None
              SpacingBefore = None
              SpacingAfter = None
              LineSpacing = None
              Indentation = None
              KeepWithNext = false
              PageBreakBefore = false
              Borders = None
              Shading = None
              TabStops = [] }
