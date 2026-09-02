namespace Kookerella.FsWordDsl

/// Page geometry - orientation, size, and margins. `HeaderFooterSet` and `SectionProperties`
/// (which pair this with a section's header/footer *content*) live in `Model.fs` instead,
/// declared together with `Block` via a mutually-recursive `and` chain, since a header or
/// footer's content is itself a `Block list` - see that file's own note on the recursion.
[<AutoOpen>]
module PageSetup =

    type PageOrientation =
        | Portrait
        | Landscape

    /// How this section begins relative to the previous one (`w:sectPr/w:type`) - Word's
    /// own `SectionMarkValues` also has `NextColumn`, only meaningful for a multi-column
    /// section, which isn't modeled distinctly here; a caller wanting that effect uses
    /// `NextPageBreak` (the same "not written, Word's own default" treatment `Writer` gives
    /// it) since this DSL's `Columns` field already covers the multi-column case itself.
    type SectionBreakType =
        | NextPageBreak
        | ContinuousBreak
        | EvenPageBreak
        | OddPageBreak

    /// A small named set covering common paper sizes, plus `OtherPaperSize` for any other
    /// OOXML `ST_PageSize` code - same "small named set + raw escape hatch" convention
    /// Excel's own `PaperSize` uses. Width/height are derived from the name at write time
    /// (swapped for `Landscape`); `Custom` gives an exact size in points for anything else.
    type PageSize =
        | Letter
        | Legal
        | A4
        | A3
        | OtherPageSize of code: int
        /// Width/height in points, portrait orientation (swapped for `Landscape` the same
        /// as the named sizes).
        | CustomPageSize of widthPoints: float * heightPoints: float

    /// All fields in points.
    type PageMargins =
        { Top: float
          Bottom: float
          Left: float
          Right: float
          Header: float
          Footer: float
          Gutter: float }

        static member Default =
            { Top = 72.0
              Bottom = 72.0
              Left = 72.0
              Right = 72.0
              Header = 36.0
              Footer = 36.0
              Gutter = 0.0 }

    /// When a footnote/endnote's own counter starts over (`w:numRestart`) - Word's own
    /// default is `ContinuousRestart` (numbered once, straight through the whole
    /// document), which is why `Footnote`/`Endnote` need no numbering settings at all by
    /// default (see `SectionProperties.FootnoteNumbering`/`EndnoteNumbering`).
    type NoteNumberRestart =
        | ContinuousRestart
        | RestartEachSection
        | RestartEachPage

    /// A section's own footnote/endnote numbering settings (`w:sectPr/w:footnotePr`/
    /// `w:endnotePr`) - `Format` reuses `Numbering.NumberFormatKind` (the same `w:numFmt`
    /// vocabulary a list level uses), though `BulletFormat` is meaningless here (Word
    /// itself never bullets a footnote/endnote reference mark) - this DSL doesn't stop a
    /// caller from setting it anyway, same "trust the caller" posture the rest of this
    /// module takes.
    type NoteNumberingSettings =
        { Format: NumberFormatKind
          StartAt: int option
          Restart: NoteNumberRestart }

        static member Default =
            { Format = DecimalFormat
              StartAt = None
              Restart = ContinuousRestart }
