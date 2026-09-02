namespace Kookerella.CsWordDsl;

/// <summary>Mirrors the F# core's own <c>TabStopAlignment</c> - five named cases plus
/// <see cref="Other"/>, a raw-OOXML-value escape hatch.</summary>
public abstract record TabStopAlignment
{
    private TabStopAlignment() { }

    public sealed record Left : TabStopAlignment;
    public sealed record Center : TabStopAlignment;
    public sealed record Right : TabStopAlignment;
    public sealed record Decimal : TabStopAlignment;
    public sealed record Bar : TabStopAlignment;
    public sealed record Other(string Raw) : TabStopAlignment;
}
