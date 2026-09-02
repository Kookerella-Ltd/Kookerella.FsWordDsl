namespace Kookerella.CsWordDsl;

/// <summary>
/// Direct/inline character formatting, written straight onto a run - never interned or
/// deduplicated (WordprocessingML has no shared stylesheet index the way SpreadsheetML
/// does). Immutable: every <c>With*</c>/<c>As*</c> method returns a new <see
/// cref="RunStyle"/>. Mirrors the F# core's own <c>RunStyle</c> record.
/// </summary>
public sealed record RunStyle
{
    public string? FontFamily { get; init; }

    /// <summary>Points.</summary>
    public double? Size { get; init; }

    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public UnderlineStyle? Underline { get; init; }
    public bool Strikethrough { get; init; }
    public Color? Color { get; init; }
    public HighlightColor? Highlight { get; init; }
    public VerticalPosition? VerticalPosition { get; init; }

    /// <summary>Renders lowercase letters as smaller uppercase ones - distinct from <see
    /// cref="AllCaps"/>, which renders every letter full-size uppercase without changing the
    /// run's own stored text.</summary>
    public bool SmallCaps { get; init; }

    public bool AllCaps { get; init; }

    /// <summary>Text present in the document but not displayed or printed until unhidden.
    /// </summary>
    public bool Hidden { get; init; }

    public static readonly RunStyle Default = new();

    public RunStyle WithFontFamily(string fontFamily) => this with { FontFamily = fontFamily };
    public RunStyle WithSize(double points) => this with { Size = points };
    public RunStyle AsBold() => this with { Bold = true };
    public RunStyle AsItalic() => this with { Italic = true };
    public RunStyle WithUnderline(UnderlineStyle underline) => this with { Underline = underline };
    public RunStyle AsStrikethrough() => this with { Strikethrough = true };
    public RunStyle WithColor(Color color) => this with { Color = color };
    public RunStyle WithHighlight(HighlightColor highlight) => this with { Highlight = highlight };
    public RunStyle WithVerticalPosition(VerticalPosition position) => this with { VerticalPosition = position };
    public RunStyle AsSmallCaps() => this with { SmallCaps = true };
    public RunStyle AsAllCaps() => this with { AllCaps = true };
    public RunStyle AsHidden() => this with { Hidden = true };
}
