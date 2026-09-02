namespace Kookerella.CsWordDsl;

/// <summary>All fields in points. <see cref="FirstLine"/>/<see cref="Hanging"/> are
/// mutually meaningful alternatives - this wrapper doesn't prevent setting both, it just
/// writes whichever are present.</summary>
public sealed record Indentation
{
    public double? Left { get; init; }
    public double? Right { get; init; }
    public double? FirstLine { get; init; }
    public double? Hanging { get; init; }

    public static readonly Indentation None = new();

    public Indentation WithLeft(double points) => this with { Left = points };
    public Indentation WithRight(double points) => this with { Right = points };
    public Indentation WithFirstLine(double points) => this with { FirstLine = points };
    public Indentation WithHanging(double points) => this with { Hanging = points };
}
