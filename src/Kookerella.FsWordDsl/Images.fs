namespace Kookerella.FsWordDsl

/// Raster images, anchored inline within a run (`Inline.Image` in `Model.fs`) - the natural
/// placement for Word, unlike Excel's cell-range-anchored `ImageEntry`. Free-floating
/// position, rotation, cropping, and formats beyond PNG/JPEG/GIF/BMP are a documented gap,
/// same posture as Excel's own `ImageEntry`.
[<AutoOpen>]
module Images =

    type ImageFormat =
        | Png
        | Jpeg
        | Gif
        | Bmp

    /// `Data` is the image file's own raw bytes (e.g. `System.IO.File.ReadAllBytes`) - this
    /// DSL doesn't decode or re-encode anything, only embeds and hands back exactly what
    /// it's given, same "opaque payload" treatment Excel's `ImageEntry.Data` gets.
    /// `WidthEmu`/`HeightEmu` are the on-page size in EMU (see `Units.fs` for conversions
    /// from points/inches/pixels) - this DSL does not read the image's own natural
    /// dimensions from its bytes, so an explicit size is always required.
    type ImageEntry =
        { Data: byte[]
          Format: ImageFormat
          WidthEmu: int64
          HeightEmu: int64
          AltText: string option }
