namespace Kookerella.FsWordDsl

/// Numbered/bulleted lists (`numbering.xml`) - `NumberingDefinition`s live on `Document`,
/// referenced from a `Paragraph` by `(numId, level)` (see `Model.fs`). WordprocessingML's
/// own numbering model separates an abstract definition from a per-document numbering
/// instance that points at one; this DSL collapses that indirection away (see
/// `Interpreter/Writer.fs`) - a caller only ever thinks in terms of one `NumberingDefinition`
/// per distinct list.
[<AutoOpen>]
module Numbering =

    /// WordprocessingML defines many more numbering formats than this names explicitly
    /// (`ordinal`, `cardinalText`, `chicago`, ...); `OtherFormat` preserves the raw OOXML
    /// `w:numFmt/@w:val` so reading and re-writing an existing document round-trips even for
    /// a format this DSL doesn't author itself - same convention as `Styles.BorderLineStyle.
    /// OtherLine`.
    type NumberFormatKind =
        /// A literal bullet glyph plus the font it renders from - Word's own bullets are
        /// conventionally drawn from a symbol font (`"Symbol"`, `"Wingdings"`, ...) rather
        /// than the paragraph's own body font, even though the glyph is stored as ordinary
        /// text (`w:lvlText`).
        | BulletFormat of glyph: char * fontFamily: string
        | DecimalFormat
        | LowerLetterFormat
        | UpperLetterFormat
        | LowerRomanFormat
        | UpperRomanFormat
        | OtherFormat of string

    /// One level of a (potentially multi-level) list definition. `Text` is the level's raw
    /// `w:lvlText` pattern - a literal glyph for `BulletFormat`, or a `"%1."`-style pattern
    /// for the numbered formats (`%1` is replaced with this level's own counter; a deeper
    /// level's pattern can reference an ancestor level's counter too, e.g. `"%1.%2"` - Core
    /// does not validate that the pattern's placeholders match the level nesting, it writes
    /// whatever text is given verbatim, same "trust the caller" posture the rest of this
    /// DSL takes on formatting fields).
    type ListLevel =
        { Format: NumberFormatKind
          Text: string
          /// Points.
          IndentLeft: float option
          /// Points.
          HangingIndent: float option
          StartAt: int option }

    /// `Id` is the number a `Paragraph`'s `Numbering` field references (see `Model.fs`) -
    /// scoped to the `Document` it lives on, not globally unique across documents.
    type NumberingDefinition = { Id: int; Levels: ListLevel list }
