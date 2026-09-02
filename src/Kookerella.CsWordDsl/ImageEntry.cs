namespace Kookerella.CsWordDsl;

public enum ImageFormat
{
    Png,
    Jpeg,
    Gif,
    Bmp
}

/// <summary>
/// A raster image, anchored inline within a run. <see cref="Data"/> is the image file's
/// own raw bytes - this wrapper doesn't decode or re-encode anything. <see
/// cref="WidthEmu"/>/<see cref="HeightEmu"/> are the on-page size in EMU (914400 per inch,
/// 12700 per point) - this wrapper does not read the image's own natural dimensions from
/// its bytes, so an explicit size is always required.
/// </summary>
public sealed record ImageEntry
{
    public required byte[] Data { get; init; }
    public required ImageFormat Format { get; init; }
    public required long WidthEmu { get; init; }
    public required long HeightEmu { get; init; }
    public string? AltText { get; init; }

    public static ImageEntry FromBytes(byte[] data, ImageFormat format, long widthEmu, long heightEmu, string? altText = null) =>
        new() { Data = data, Format = format, WidthEmu = widthEmu, HeightEmu = heightEmu, AltText = altText };

    /// <summary>Sizes from inches rather than raw EMU.</summary>
    public static ImageEntry FromBytesInches(byte[] data, ImageFormat format, double widthInches, double heightInches, string? altText = null) =>
        FromBytes(data, format, (long)Math.Round(widthInches * 914400.0), (long)Math.Round(heightInches * 914400.0), altText);
}
