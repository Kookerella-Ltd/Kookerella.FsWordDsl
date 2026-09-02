namespace Kookerella.CsWordDsl;

/// <summary>
/// Direct/inline paragraph formatting - written straight onto the paragraph. <see
/// cref="Paragraph.StyleId"/> supplies the named-style layer; this is the direct-formatting
/// layer on top of it. Mirrors the F# core's own <c>ParagraphFormat</c> record.
/// </summary>
public sealed record ParagraphFormat
{
    public ParagraphAlignment? Alignment { get; init; }
    public double? SpacingBefore { get; init; }
    public double? SpacingAfter { get; init; }
    public LineSpacingRule? LineSpacing { get; init; }
    public Indentation? Indentation { get; init; }
    public bool KeepWithNext { get; init; }
    public bool PageBreakBefore { get; init; }

    /// <summary>The paragraph's own border box - independent of any table border the
    /// paragraph might also sit inside.</summary>
    public BorderStyle? Borders { get; init; }

    /// <summary>Background fill behind the paragraph's text.</summary>
    public Color? Shading { get; init; }

    /// <summary>Custom tab stops - an empty list means "no custom tabs", not "clear
    /// Word's own default tab stops every half-inch".</summary>
    public IReadOnlyList<TabStop> TabStops { get; init; } = Array.Empty<TabStop>();

    public static readonly ParagraphFormat Default = new();

    public ParagraphFormat WithAlignment(ParagraphAlignment alignment) => this with { Alignment = alignment };
    public ParagraphFormat WithSpacingBefore(double points) => this with { SpacingBefore = points };
    public ParagraphFormat WithSpacingAfter(double points) => this with { SpacingAfter = points };
    public ParagraphFormat WithLineSpacing(LineSpacingRule rule) => this with { LineSpacing = rule };
    public ParagraphFormat WithIndentation(Indentation indentation) => this with { Indentation = indentation };
    public ParagraphFormat AsKeepWithNext() => this with { KeepWithNext = true };
    public ParagraphFormat AsPageBreakBefore() => this with { PageBreakBefore = true };
    public ParagraphFormat WithBorders(BorderStyle borders) => this with { Borders = borders };
    public ParagraphFormat WithShading(Color color) => this with { Shading = color };
    public ParagraphFormat WithTabStops(params TabStop[] tabStops) => this with { TabStops = tabStops };
}
