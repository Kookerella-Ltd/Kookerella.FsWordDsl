namespace Kookerella.CsWordDsl;

/// <summary>
/// A small catalog of the style ids real Word documents reach for constantly, mirroring the
/// F# core's own <c>BuiltInStyles</c> module. Not exhaustive - any other <see
/// cref="StyleDefinition"/> works the same way; these are just the common case pre-built.
/// </summary>
public static class BuiltInStyles
{
    public static readonly StyleDefinition Normal = new()
    {
        Id = "Normal",
        Name = "Normal",
        Type = StyleTargetType.Paragraph,
        RunFormat = RunStyle.Default.WithFontFamily("Calibri").WithSize(11.0)
    };

    public static readonly StyleDefinition Heading1 = new()
    {
        Id = "Heading1",
        Name = "heading 1",
        Type = StyleTargetType.Paragraph,
        BasedOn = "Normal",
        RunFormat = RunStyle.Default with { Bold = true, Size = 16.0, Color = new Color.Rgb(47, 84, 150) },
        ParaFormat = ParagraphFormat.Default with { SpacingBefore = 12.0, SpacingAfter = 6.0, KeepWithNext = true }
    };

    public static readonly StyleDefinition Heading2 = new()
    {
        Id = "Heading2",
        Name = "heading 2",
        Type = StyleTargetType.Paragraph,
        BasedOn = "Normal",
        RunFormat = RunStyle.Default with { Bold = true, Size = 14.0, Color = new Color.Rgb(47, 84, 150) },
        ParaFormat = ParagraphFormat.Default with { SpacingBefore = 10.0, SpacingAfter = 4.0, KeepWithNext = true }
    };

    public static readonly StyleDefinition Heading3 = new()
    {
        Id = "Heading3",
        Name = "heading 3",
        Type = StyleTargetType.Paragraph,
        BasedOn = "Normal",
        RunFormat = RunStyle.Default with { Bold = true, Size = 13.0, Color = new Color.Rgb(47, 84, 150) },
        ParaFormat = ParagraphFormat.Default with { SpacingBefore = 8.0, SpacingAfter = 4.0, KeepWithNext = true }
    };

    public static readonly StyleDefinition Title = new()
    {
        Id = "Title",
        Name = "Title",
        Type = StyleTargetType.Paragraph,
        BasedOn = "Normal",
        RunFormat = RunStyle.Default with { Bold = true, Size = 28.0 },
        ParaFormat = ParagraphFormat.Default with { SpacingAfter = 12.0 }
    };

    public static readonly StyleDefinition ListParagraph = new()
    {
        Id = "ListParagraph",
        Name = "List Paragraph",
        Type = StyleTargetType.Paragraph,
        BasedOn = "Normal",
        ParaFormat = ParagraphFormat.Default with { Indentation = Indentation.None with { Left = 36.0 } }
    };

    /// <summary>The character style Word applies to a hyperlink's own runs - <see
    /// cref="Inline"/>'s own <c>Hyperlink</c> factory applies this automatically.</summary>
    public static readonly StyleDefinition HyperlinkCharStyle = new()
    {
        Id = "Hyperlink",
        Name = "Hyperlink",
        Type = StyleTargetType.Character,
        RunFormat = RunStyle.Default with { Color = new Color.Rgb(5, 99, 193), Underline = new UnderlineStyle.Single() }
    };

    public static readonly StyleDefinition FootnoteReferenceCharStyle = new()
    {
        Id = "FootnoteReference",
        Name = "footnote reference",
        Type = StyleTargetType.Character,
        RunFormat = RunStyle.Default with { VerticalPosition = VerticalPosition.Superscript }
    };

    public static readonly StyleDefinition EndnoteReferenceCharStyle = new()
    {
        Id = "EndnoteReference",
        Name = "endnote reference",
        Type = StyleTargetType.Character,
        RunFormat = RunStyle.Default with { VerticalPosition = VerticalPosition.Superscript }
    };

    public static readonly StyleDefinition FootnoteTextStyle = new()
    {
        Id = "FootnoteText",
        Name = "footnote text",
        Type = StyleTargetType.Paragraph,
        BasedOn = "Normal",
        RunFormat = RunStyle.Default with { Size = 10.0 }
    };

    public static readonly StyleDefinition EndnoteTextStyle = new()
    {
        Id = "EndnoteText",
        Name = "endnote text",
        Type = StyleTargetType.Paragraph,
        BasedOn = "Normal",
        RunFormat = RunStyle.Default with { Size = 10.0 }
    };

    /// <summary>Every built-in above - the C# analog of the F# core's <c>BuiltInStyles.
    /// all</c>, used as <see cref="Document.Styles"/>'s own default.</summary>
    public static readonly IReadOnlyList<StyleDefinition> All =
    [
        Normal, Heading1, Heading2, Heading3, Title, ListParagraph,
        HyperlinkCharStyle, FootnoteReferenceCharStyle, EndnoteReferenceCharStyle,
        FootnoteTextStyle, EndnoteTextStyle
    ];
}
