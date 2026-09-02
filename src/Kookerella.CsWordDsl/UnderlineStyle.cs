namespace Kookerella.CsWordDsl;

/// <summary>Mirrors the F# core's own <c>UnderlineStyle</c> - six named cases plus <see
/// cref="Other"/>, a raw-OOXML-value escape hatch so reading and re-writing an existing
/// document round-trips even for an underline kind this wrapper doesn't name explicitly.
/// </summary>
public abstract record UnderlineStyle
{
    private UnderlineStyle() { }

    public sealed record Single : UnderlineStyle;
    public sealed record Double : UnderlineStyle;
    public sealed record Thick : UnderlineStyle;
    public sealed record Dotted : UnderlineStyle;
    public sealed record Dashed : UnderlineStyle;
    public sealed record Wavy : UnderlineStyle;
    public sealed record Other(string Raw) : UnderlineStyle;
}
