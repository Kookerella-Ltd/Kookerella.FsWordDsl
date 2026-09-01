namespace Kookerella.FsWordDsl

/// Conversions between the physical units WordprocessingML uses on the wire and plain
/// points/inches, the units this DSL exposes to callers. There is no single "natural" unit
/// the way Excel's column-width-in-character-widths already matches OOXML directly - Word
/// mixes twentieths-of-a-point ("twips", page sizes/margins/indentation/spacing) and
/// English Metric Units ("EMU", DrawingML image sizing) depending on the part of the
/// schema, so this module is the one place both live.
module Units =

    /// Twips ("twentieths of a point") - WordprocessingML's unit for page size, margins,
    /// indentation, and spacing (`w:pgSz`, `w:pgMar`, `w:ind`, `w:spacing`, ...).
    let pointsToTwips (points: float) : int = int (System.Math.Round(points * 20.0))

    let twipsToPoints (twips: int) : float = float twips / 20.0

    let inchesToTwips (inches: float) : int = pointsToTwips (inches * 72.0)

    let twipsToInches (twips: int) : float = twipsToPoints twips / 72.0

    /// EMU ("English Metric Units") - DrawingML's unit for image/shape sizing, 914400 per
    /// inch and 12700 per point. `ImageEntry.WidthEmu`/`HeightEmu` are already in this unit
    /// since that's what a caller placing an image at an exact on-page size needs, but these
    /// helpers make it easy to compute from more familiar units instead.
    let inchesToEmu (inches: float) : int64 = int64 (System.Math.Round(inches * 914400.0))

    let emuToInches (emu: int64) : float = float emu / 914400.0

    let pointsToEmu (points: float) : int64 = int64 (System.Math.Round(points * 12700.0))

    let emuToPoints (emu: int64) : float = float emu / 12700.0

    /// Pixels at the conventional 96 DPI screen resolution - handy for sizing an image
    /// against its own natural pixel dimensions rather than an explicit physical size.
    let pixelsToEmu (pixels: int) : int64 = inchesToEmu (float pixels / 96.0)
