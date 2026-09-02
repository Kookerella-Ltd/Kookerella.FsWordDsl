namespace Kookerella.CsWordDsl;

/// <summary>Mirrors the F# core's own <c>BorderLineStyle</c> - six named cases plus <see
/// cref="Other"/>, a raw-OOXML-value escape hatch.</summary>
public abstract record BorderLineStyle
{
    private BorderLineStyle() { }

    public sealed record Single : BorderLineStyle;
    public sealed record Thick : BorderLineStyle;
    public sealed record Double : BorderLineStyle;
    public sealed record Dotted : BorderLineStyle;
    public sealed record Dashed : BorderLineStyle;
    public sealed record Wave : BorderLineStyle;
    public sealed record Other(string Raw) : BorderLineStyle;
}
