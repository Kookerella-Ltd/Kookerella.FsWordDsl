namespace Kookerella.CsWordDsl;

/// <summary>All fields in points.</summary>
public sealed record PageMargins
{
    public double Top { get; init; } = 72.0;
    public double Bottom { get; init; } = 72.0;
    public double Left { get; init; } = 72.0;
    public double Right { get; init; } = 72.0;
    public double Header { get; init; } = 36.0;
    public double Footer { get; init; } = 36.0;
    public double Gutter { get; init; }

    public static readonly PageMargins Default = new();
}
