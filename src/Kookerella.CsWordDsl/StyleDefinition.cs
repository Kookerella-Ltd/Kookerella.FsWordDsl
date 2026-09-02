namespace Kookerella.CsWordDsl;

public enum StyleTargetType
{
    Paragraph,
    Character
}

/// <summary>
/// A named style (<c>styles.xml</c>) - the one styling concept central to how real Word
/// documents are authored (a <see cref="Paragraph.StyleId"/> referencing <c>"Heading1"</c>
/// is far more common in practice than direct formatting on every paragraph). This wrapper
/// does not resolve the <see cref="BasedOn"/> inheritance chain itself - that's real Word's
/// job when it renders/edits the file.
/// </summary>
public sealed record StyleDefinition
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public required StyleTargetType Type { get; init; }
    public string? BasedOn { get; init; }
    public RunStyle? RunFormat { get; init; }
    public ParagraphFormat? ParaFormat { get; init; }

    public StyleDefinition WithBasedOn(string basedOnId) => this with { BasedOn = basedOnId };
    public StyleDefinition WithRunFormat(RunStyle format) => this with { RunFormat = format };
    public StyleDefinition WithParaFormat(ParagraphFormat format) => this with { ParaFormat = format };
}
