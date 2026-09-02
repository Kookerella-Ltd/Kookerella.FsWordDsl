namespace Kookerella.CsWordDsl;

/// <summary>When a footnote/endnote's own counter starts over - Word's own default is
/// <see cref="Continuous"/> (numbered once, straight through the whole document).</summary>
public enum NoteNumberRestart
{
    Continuous,
    EachSection,
    EachPage
}

/// <summary>A section's own footnote/endnote numbering settings. <see cref="Format"/>
/// reuses <see cref="NumberFormatKind"/> (<see cref="NumberFormatKind.Bullet"/> is
/// meaningless here, but this wrapper doesn't stop a caller setting it anyway).</summary>
public sealed record NoteNumberingSettings
{
    public required NumberFormatKind Format { get; init; }
    public int? StartAt { get; init; }
    public NoteNumberRestart Restart { get; init; } = NoteNumberRestart.Continuous;

    public static readonly NoteNumberingSettings Default = new() { Format = new NumberFormatKind.Decimal() };
}
