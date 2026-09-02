namespace Kookerella.CsWordDsl;

/// <summary>A small named set covering common paper sizes, plus <see cref="Other"/>/<see
/// cref="Custom"/> escape hatches. Width/height are derived from the name at write time
/// (swapped for landscape orientation).</summary>
public abstract record PageSize
{
    private PageSize() { }

    public sealed record Letter : PageSize;
    public sealed record Legal : PageSize;
    public sealed record A4 : PageSize;
    public sealed record A3 : PageSize;

    /// <summary>Any other raw OOXML <c>ST_PageSize</c> code.</summary>
    public sealed record Other(int Code) : PageSize;

    /// <summary>Width/height in points, portrait orientation (swapped for landscape the
    /// same as the named sizes).</summary>
    public sealed record Custom(double WidthPoints, double HeightPoints) : PageSize;
}
