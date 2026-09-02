namespace Kookerella.CsWordDsl;

/// <summary>Mirrors the F# core's own <c>LineSpacingRule</c>.</summary>
public abstract record LineSpacingRule
{
    private LineSpacingRule() { }

    public sealed record Single : LineSpacingRule;
    public sealed record OnePointFive : LineSpacingRule;
    public sealed record DoubleSpacing : LineSpacingRule;

    /// <summary>Points - line height is at least this tall, growing to fit taller content.</summary>
    public sealed record AtLeast(double Points) : LineSpacingRule;

    /// <summary>Points - line height is fixed exactly, regardless of content.</summary>
    public sealed record Exactly(double Points) : LineSpacingRule;

    /// <summary>A multiple of single line spacing, e.g. 1.15.</summary>
    public sealed record Multiple(double Factor) : LineSpacingRule;
}
