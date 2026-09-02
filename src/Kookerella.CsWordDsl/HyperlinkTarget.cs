namespace Kookerella.CsWordDsl;

/// <summary>A hyperlink's destination - external (any URL, including <c>mailto:</c>) or
/// internal (a same-document bookmark reference).</summary>
public abstract record HyperlinkTarget
{
    private HyperlinkTarget() { }

    public sealed record ExternalUrl(string Url) : HyperlinkTarget;
    public sealed record InternalBookmark(string Name) : HyperlinkTarget;
}
