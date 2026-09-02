namespace Kookerella.CsWordDsl;

/// <summary>Mirrors the F# core's own <c>NumberFormatKind</c> - five named cases plus
/// <see cref="Other"/>, a raw-OOXML-value escape hatch.</summary>
public abstract record NumberFormatKind
{
    private NumberFormatKind() { }

    /// <summary>A literal bullet glyph plus the font it renders from - Word's own bullets
    /// are conventionally drawn from a symbol font (<c>"Symbol"</c>, <c>"Wingdings"</c>,
    /// ...) rather than the paragraph's own body font.</summary>
    public sealed record Bullet(char Glyph, string FontFamily) : NumberFormatKind;

    public sealed record Decimal : NumberFormatKind;
    public sealed record LowerLetter : NumberFormatKind;
    public sealed record UpperLetter : NumberFormatKind;
    public sealed record LowerRoman : NumberFormatKind;
    public sealed record UpperRoman : NumberFormatKind;
    public sealed record Other(string Raw) : NumberFormatKind;
}
