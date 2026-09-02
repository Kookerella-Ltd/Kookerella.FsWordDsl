namespace Kookerella.CsWordDsl;

/// <summary>
/// Document-level metadata (<c>docProps/core.xml</c> and <c>docProps/app.xml</c>). All
/// fields optional - a document with no properties set round-trips back to <see
/// cref="Default"/> exactly.
/// </summary>
public sealed record DocumentProperties
{
    public string? Title { get; init; }
    public string? Author { get; init; }
    public string? Subject { get; init; }
    public string? Keywords { get; init; }

    /// <summary>Word's own UI now calls this field "Comments" even though OOXML's own
    /// package-level name for it is <c>dc:description</c>.</summary>
    public string? Comments { get; init; }

    public string? Category { get; init; }
    public string? Company { get; init; }

    public static readonly DocumentProperties Default = new();

    public DocumentProperties WithTitle(string title) => this with { Title = title };
    public DocumentProperties WithAuthor(string author) => this with { Author = author };
    public DocumentProperties WithSubject(string subject) => this with { Subject = subject };
    public DocumentProperties WithKeywords(string keywords) => this with { Keywords = keywords };
    public DocumentProperties WithComments(string comments) => this with { Comments = comments };
    public DocumentProperties WithCategory(string category) => this with { Category = category };
    public DocumentProperties WithCompany(string company) => this with { Company = company };
}
