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
