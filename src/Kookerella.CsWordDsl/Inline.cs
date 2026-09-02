namespace Kookerella.CsWordDsl;

/// <summary>
/// A paragraph's inline content - a closed set of immutable cases, mirroring the F# core's
/// own <c>Inline</c> discriminated union (same "sealed hierarchy with a private base
/// constructor" pattern every closed-case type in this wrapper uses). Unlike a single
/// uniformly-styled cell value, a <see cref="Paragraph"/>'s <see cref="Paragraph.Inlines"/>
/// naturally hold several independently-styled runs, so rich text is first-class here.
/// </summary>
public abstract record Inline
{
    private Inline() { }

    /// <summary><paramref name="StyleId"/> references a character style (e.g.
    /// <c>"Hyperlink"</c>); <paramref name="Style"/> is direct formatting layered on top -
    /// either, both, or neither may be given.</summary>
    public sealed record Run(string Text, RunStyle? Style = null, string? StyleId = null) : Inline;

    public sealed record LineBreak : Inline;
    public sealed record Tab : Inline;

    /// <summary>An explicit manual page break, mid-paragraph - distinct from <see
    /// cref="ParagraphFormat.PageBreakBefore"/>, which breaks before an entire paragraph.
    /// </summary>
    public sealed record PageBreak : Inline;

    public sealed record Image(ImageEntry Entry) : Inline;

    public sealed record Hyperlink(HyperlinkTarget Target, IReadOnlyList<Inline> Runs, string? Tooltip = null) : Inline;

    /// <summary>Scoped to within a single paragraph - the ergonomic, common case. A
    /// bookmark spanning more than one paragraph is <see cref="BookmarkRangeStart"/>/<see
    /// cref="BookmarkRangeEnd"/> instead.</summary>
    public sealed record Bookmark(string Name, IReadOnlyList<Inline> Content) : Inline;

    /// <summary>A bookmark boundary marker usable on its own, for a bookmark spanning more
    /// than one paragraph - <paramref name="Name"/> must match its corresponding <see
    /// cref="BookmarkRangeEnd"/> elsewhere in the same document.</summary>
    public sealed record BookmarkRangeStart(string Name) : Inline;

    public sealed record BookmarkRangeEnd(string Name) : Inline;

    /// <summary><paramref name="Date"/> = <see langword="null"/> is written as "now" at
    /// write time. Scoped to within a single paragraph - the ergonomic, common case. A
    /// comment spanning more than one paragraph is <see cref="CommentRangeStart"/>/<see
    /// cref="CommentRangeEnd"/> instead.</summary>
    public sealed record Comment(string Author, string? Initials, DateTime? Date, string Text, IReadOnlyList<Inline> Content) : Inline;

    /// <summary>A comment boundary marker usable on its own, for a comment spanning more
    /// than one paragraph. <paramref name="Id"/> is a caller-chosen correlation key
    /// matching this start to its <see cref="CommentRangeEnd"/> elsewhere in the document -
    /// it is write-time-only and does not round-trip (OOXML has nowhere to persist an
    /// arbitrary string alongside a comment range, only a numeric id the writer assigns).
    /// </summary>
    public sealed record CommentRangeStart(string Id, string Author, string? Initials, DateTime? Date, string Text) : Inline;

    public sealed record CommentRangeEnd(string Id) : Inline;

    /// <summary>A "simple field" - raw field instruction text (e.g. <c>"PAGE"</c>) plus the
    /// cached display text Word showed before its own recalculation. This wrapper never
    /// evaluates a field itself - real Word recalculates on open and overwrites <paramref
    /// name="CachedResult"/>.</summary>
    public sealed record Field(string Instruction, string? CachedResult = null) : Inline;

    /// <summary>A footnote/endnote reference mark - <paramref name="Content"/> is the
    /// note's own body. The writer prepends the note-reference-mark run to the body's
    /// first paragraph automatically.</summary>
    public sealed record Footnote(IReadOnlyList<Block> Content) : Inline;

    public sealed record Endnote(IReadOnlyList<Block> Content) : Inline;

    /// <summary>Track changes - wraps arbitrary inline content, marking it as inserted or
    /// deleted by <paramref name="Revision"/>.</summary>
    public sealed record TrackedChange(Revision Revision, IReadOnlyList<Inline> Content) : Inline;

    /// <summary>A structured document tag (content control) sitting inside a single
    /// paragraph. The block-level counterpart is <see cref="Block.ContentControlBlock"/>.
    /// </summary>
    public sealed record ContentControl(ContentControlProps Props, IReadOnlyList<Inline> Content) : Inline;

    // ----- Convenience factories for the highest-frequency cases, matching the F# core's
    // own DocumentDsl sugar. Still spelled with `new` for the plain cases above - a nested
    // case type and a same-named static factory can't coexist in C#, so these are only for
    // shapes that genuinely differ from (rather than just rename) their underlying case.

    /// <summary>A hyperlink over plain text - applies <see
    /// cref="BuiltInStyles.HyperlinkCharStyle"/> automatically.</summary>
    public static Inline HyperlinkText(string text, HyperlinkTarget target, string? tooltip = null) =>
        new Hyperlink(target, [new Run(text, StyleId: BuiltInStyles.HyperlinkCharStyle.Id)], tooltip);

    /// <summary>A footnote over plain text - applies <see
    /// cref="BuiltInStyles.FootnoteTextStyle"/> to the note body's own paragraph.</summary>
    public static Inline FootnoteText(string text) =>
        new Footnote([new Block.ParagraphBlock(new Paragraph { Inlines = [new Run(text)], StyleId = BuiltInStyles.FootnoteTextStyle.Id })]);

    public static Inline EndnoteText(string text) =>
        new Endnote([new Block.ParagraphBlock(new Paragraph { Inlines = [new Run(text)], StyleId = BuiltInStyles.EndnoteTextStyle.Id })]);

    /// <summary>Marks <paramref name="content"/> as inserted under track changes.
    /// <paramref name="date"/> defaults to "now" at write time when omitted.</summary>
    public static Inline Inserted(IReadOnlyList<Inline> content, string author, DateTime? date = null) =>
        new TrackedChange(new Revision(RevisionKind.Inserted, author, date), content);

    /// <summary>Marks <paramref name="content"/> as deleted under track changes.</summary>
    public static Inline Deleted(IReadOnlyList<Inline> content, string author, DateTime? date = null) =>
        new TrackedChange(new Revision(RevisionKind.Deleted, author, date), content);
}
