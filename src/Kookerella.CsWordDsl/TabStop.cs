namespace Kookerella.CsWordDsl;

/// <summary>One custom tab stop. <see cref="Position"/> is in points, measured from the
/// left margin.</summary>
public sealed record TabStop(double Position, TabStopAlignment Alignment, TabLeader Leader = TabLeader.None);
