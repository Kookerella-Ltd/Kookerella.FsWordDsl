namespace Kookerella.CsWordDsl;

/// <summary>Which <c>w:lock</c> value restricts editing a content control - <see
/// langword="null"/> on <see cref="ContentControlProps.Lock"/> is Word's own default
/// (unlocked, not written).</summary>
public enum ContentControlLock
{
    /// <summary>The control itself can't be deleted, but its content can still be edited.</summary>
    LockDeletion,

    /// <summary>The control's content can't be edited, but the control itself can still be
    /// deleted.</summary>
    LockContentEditing,

    /// <summary>Neither the control nor its content can be edited.</summary>
    LockDeletionAndContentEditing
}

/// <summary><see cref="Alias"/>/<see cref="Tag"/> are Word's own human-readable
/// title/machine-readable id - both optional.</summary>
public sealed record ContentControlProps
{
    public string? Alias { get; init; }
    public string? Tag { get; init; }
    public ContentControlLock? Lock { get; init; }
    public required ContentControlType Type { get; init; }

    public ContentControlProps WithAlias(string alias) => this with { Alias = alias };
    public ContentControlProps WithTag(string tag) => this with { Tag = tag };
    public ContentControlProps WithLock(ContentControlLock @lock) => this with { Lock = @lock };
}
