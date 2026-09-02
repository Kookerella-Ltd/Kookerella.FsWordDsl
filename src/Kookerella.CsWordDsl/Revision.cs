namespace Kookerella.CsWordDsl;

public enum RevisionKind
{
    Inserted,
    Deleted
}

/// <summary>Track-changes metadata - who made an insertion/deletion and when. <see
/// cref="Date"/> defaults to "now" at write time when omitted.</summary>
public sealed record Revision(RevisionKind Kind, string Author, DateTime? Date = null);
