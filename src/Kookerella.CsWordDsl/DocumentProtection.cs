namespace Kookerella.CsWordDsl;

/// <summary>Which single kind of edit Word still allows while the document is protected -
/// these are mutually exclusive in real Word.</summary>
public enum EditRestriction
{
    ReadOnly,
    CommentsOnly,
    TrackedChangesOnly,
    FormsOnly
}

/// <summary>Document-level editing restrictions - Word has no per-section equivalent of a
/// spreadsheet's per-sheet protection, only this one document-wide setting. <see
/// cref="Password"/> is hashed with the modern salted-iterated-SHA512 scheme and never
/// round-trips back to plaintext.</summary>
public sealed record DocumentProtection
{
    public EditRestriction? Edit { get; init; }
    public string? Password { get; init; }

    public static DocumentProtection With(EditRestriction edit, string? password = null) => new() { Edit = edit, Password = password };
}
