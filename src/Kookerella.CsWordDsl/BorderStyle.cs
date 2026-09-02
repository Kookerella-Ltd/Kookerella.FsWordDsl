namespace Kookerella.CsWordDsl;

/// <summary>Reused for both paragraph borders and table/cell borders. Diagonal cell
/// borders and a paragraph's <c>between</c>/<c>bar</c> sides aren't modeled - see the F#
/// core's own <c>BorderStyle</c> doc comment.</summary>
public sealed record BorderStyle
{
    public BorderSide? Left { get; init; }
    public BorderSide? Right { get; init; }
    public BorderSide? Top { get; init; }
    public BorderSide? Bottom { get; init; }

    public static readonly BorderStyle None = new();

    public BorderStyle WithLeft(BorderSide side) => this with { Left = side };
    public BorderStyle WithRight(BorderSide side) => this with { Right = side };
    public BorderStyle WithTop(BorderSide side) => this with { Top = side };
    public BorderStyle WithBottom(BorderSide side) => this with { Bottom = side };

    /// <summary>Sets all four edges to the same side in one call.</summary>
    public BorderStyle WithAllSides(BorderSide side) => this with { Left = side, Right = side, Top = side, Bottom = side };
}
