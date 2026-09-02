namespace Kookerella.CsWordDsl;

/// <summary>
/// Which of Word's content-control kinds this is, plus that kind's own extra data.
/// Mirrors the F# core's own <c>ContentControlType</c>. See <c>Inline</c>/<c>Block</c>'s
/// own <c>ContentControl</c>/<c>ContentControlBlock</c> factories for how this plugs in.
/// </summary>
public abstract record ContentControlType
{
    private ContentControlType() { }

    public sealed record PlainText(bool MultiLine = false) : ContentControlType;

    public sealed record RichText : ContentControlType;

    /// <summary><paramref name="Editable"/> distinguishes a dropdown list (pick only,
    /// <see langword="false"/>) from a combo box (pick or type free text, <see
    /// langword="true"/>) - both carry the identical <paramref name="Items"/> shape.
    /// </summary>
    public sealed record DropDown(IReadOnlyList<(string DisplayText, string Value)> Items, bool Editable = false) : ContentControlType;

    /// <summary>The control's own currently-displayed text still lives in the wrapping
    /// <c>Inline</c>/<c>Block</c> case's own content - <paramref name="FullDate"/>/
    /// <paramref name="Format"/> here are metadata about how that text was produced, not
    /// the text itself.</summary>
    public sealed record Date(DateTime? FullDate = null, string? Format = null) : ContentControlType;

    /// <summary><paramref name="CheckedSymbol"/>/<paramref name="UncheckedSymbol"/> are
    /// optional custom checked/unchecked glyphs, each a (font, hex character code) pair,
    /// e.g. <c>("Wingdings", "2612")</c>.</summary>
    public sealed record CheckBox(
        bool Checked,
        (string Font, string Code)? CheckedSymbol = null,
        (string Font, string Code)? UncheckedSymbol = null
    ) : ContentControlType;
}
