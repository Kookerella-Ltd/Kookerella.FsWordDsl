namespace Kookerella.CsWordDsl;

/// <summary>
/// A run/shading/border color - a closed set of immutable cases, mirroring the F# core's
/// own <c>Color</c> discriminated union (same "sealed hierarchy with a private base
/// constructor" pattern every closed-case type in this wrapper uses).
/// </summary>
public abstract record Color
{
    private Color() { }

    public sealed record Rgb(byte R, byte G, byte B) : Color;

    public sealed record Auto : Color;

    /// <summary><paramref name="Fallback"/> is always written alongside the theme token -
    /// see the F# core's own <c>Color.Theme</c> doc comment for why (this DSL has no theme
    /// part to resolve <paramref name="Kind"/> against, so real Word does the resolving).
    /// <paramref name="Tint"/>/<paramref name="Shade"/> are 0.0-1.0, stored on the wire as a
    /// single byte, so an arbitrary value round-trips to the nearest 1/255, not bit-for-bit.
    /// </summary>
    public sealed record Theme(ThemeColorKind Kind, Rgb Fallback, double? Tint = null, double? Shade = null) : Color;

    // Convenience constants, mirroring the F# core's own `Color` module values.
    public static readonly Color Black = new Rgb(0, 0, 0);
    public static readonly Color White = new Rgb(255, 255, 255);
    public static readonly Color Red = new Rgb(255, 0, 0);
    public static readonly Color Green = new Rgb(0, 128, 0);
    public static readonly Color Blue = new Rgb(0, 0, 255);
    public static readonly Color Yellow = new Rgb(255, 255, 0);
}
