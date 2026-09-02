namespace Kookerella.CsWordDsl;

/// <summary>
/// One paragraph's content plus its formatting. <see cref="StyleId"/> references a named
/// style - the inheritance layer; <see cref="Format"/> is direct/inline formatting layered
/// on top. <see cref="Numbering"/> = (numId, level) places this paragraph in a
/// numbered/bulleted list.
/// </summary>
public sealed record Paragraph
{
    public IReadOnlyList<Inline> Inlines { get; init; } = Array.Empty<Inline>();
    public string? StyleId { get; init; }
    public ParagraphFormat? Format { get; init; }
    public (int NumId, int Level)? Numbering { get; init; }

    /// <summary>Whether this paragraph's own closing mark (the boundary to the next
    /// paragraph) was itself inserted or deleted under track changes - distinct from any
    /// <see cref="Inline.TrackedChange"/> wrapping the paragraph's own <see
    /// cref="Inlines"/>, which marks the content rather than the mark.</summary>
    public Revision? MarkRevision { get; init; }

    public Paragraph WithStyleId(string styleId) => this with { StyleId = styleId };
    public Paragraph WithFormat(ParagraphFormat format) => this with { Format = format };
    public Paragraph WithNumbering(int numId, int level) => this with { Numbering = (numId, level) };
    public Paragraph WithMarkRevision(Revision revision) => this with { MarkRevision = revision };
}
