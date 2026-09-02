namespace Kookerella.CsWordDsl;

/// <summary>
/// One section's page setup - a document is a sequence of <see cref="Section"/>s, each
/// with its own. <see cref="BreakType"/> is how this section begins relative to the
/// previous one - meaningless, and not written, for the very first section. <see
/// cref="Columns"/> is the number of equal-width text columns (1 = the ordinary
/// single-column case).
/// </summary>
public sealed record SectionProperties
{
    public PageSize PageSize { get; init; } = new PageSize.Letter();
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;
    public PageMargins Margins { get; init; } = PageMargins.Default;
    public HeaderFooterSet? Header { get; init; }
    public HeaderFooterSet? Footer { get; init; }
    public int? PageNumberStart { get; init; }
    public int Columns { get; init; } = 1;
    public SectionBreakType BreakType { get; init; } = SectionBreakType.NextPage;

    /// <summary><see langword="null"/> is Word's own default (continuous decimal
    /// numbering starting at 1, straight through the document).</summary>
    public NoteNumberingSettings? FootnoteNumbering { get; init; }
    public NoteNumberingSettings? EndnoteNumbering { get; init; }

    public static readonly SectionProperties Default = new();

    public SectionProperties WithPageSize(PageSize size) => this with { PageSize = size };
    public SectionProperties WithOrientation(PageOrientation orientation) => this with { Orientation = orientation };
    public SectionProperties WithMargins(PageMargins margins) => this with { Margins = margins };
    public SectionProperties WithHeader(HeaderFooterSet header) => this with { Header = header };
    public SectionProperties WithFooter(HeaderFooterSet footer) => this with { Footer = footer };
    public SectionProperties WithPageNumberStart(int start) => this with { PageNumberStart = start };
    public SectionProperties WithColumns(int columns) => this with { Columns = columns };
    public SectionProperties WithBreakType(SectionBreakType breakType) => this with { BreakType = breakType };
    public SectionProperties WithFootnoteNumbering(NoteNumberingSettings settings) => this with { FootnoteNumbering = settings };
    public SectionProperties WithEndnoteNumbering(NoteNumberingSettings settings) => this with { EndnoteNumbering = settings };
}

public sealed record Section(IReadOnlyList<Block> Body, SectionProperties Properties)
{
    public static Section Of(IReadOnlyList<Block> body) => new(body, SectionProperties.Default);
    public static Section With(SectionProperties properties, IReadOnlyList<Block> body) => new(body, properties);
}
