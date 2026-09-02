namespace Kookerella.CsWordDsl;

/// <summary>
/// A Word document - an ordered list of sections, named styles, numbering definitions,
/// optional document-level protection, plus an optional VBA project. Pure data, same as
/// the F# core's own <c>Document</c> - deliberately has no <c>Save</c>/<c>Load</c> methods
/// on it, since those are I/O (see <see cref="DocumentIO"/>, the one place this wrapper
/// does anything side-effecting). Immutable - every <c>With*</c>/<c>Add*</c> method returns
/// a new <see cref="Document"/> rather than mutating in place.
/// </summary>
public sealed record Document
{
    public IReadOnlyList<Section> Sections { get; init; } = Array.Empty<Section>();

    /// <summary>Defaults to <see cref="BuiltInStyles.All"/> so <c>StyleId = "Heading1"</c>
    /// (or any other built-in id) just works without registering it first - use <see
    /// cref="WithStyles"/> to replace or extend that set.</summary>
    public IReadOnlyList<StyleDefinition> Styles { get; init; } = BuiltInStyles.All;

    public IReadOnlyList<NumberingDefinition> Numbering { get; init; } = Array.Empty<NumberingDefinition>();
    public DocumentProtection? Protection { get; init; }

    private readonly byte[]? _vbaProject;

    /// <summary>A macro-enabled template's VBA project, as the raw bytes of a
    /// <c>word/vbaProject.bin</c> - a compiled OLE/CFBF binary, not source text. Nothing in
    /// this stack parses, generates, or edits VBA; the bytes are embedded and handed back
    /// verbatim. Set it via <see cref="WithVbaProject"/> and save to a
    /// <c>.docm</c>/<c>.dotm</c> path (the file's content type switches to macro-enabled
    /// automatically, but real Word expects the extension to match before it will trust
    /// and run macros).</summary>
    public byte[]? VbaProject
    {
        get => _vbaProject;
        init => _vbaProject = value;
    }

    /// <summary>Title/Author/Subject/etc.</summary>
    public DocumentProperties Properties { get; init; } = DocumentProperties.Default;

    /// <summary>Custom table style definitions - referenced from a <see
    /// cref="TableEntry.Style"/>'s own <see cref="TableStyleRef.Name"/> the same way a
    /// built-in name is.</summary>
    public IReadOnlyList<TableStyleDefinition> TableStyles { get; init; } = Array.Empty<TableStyleDefinition>();

    public static Document Create(params Section[] sections) => new() { Sections = sections };

    public Document AddSection(Section section) => this with { Sections = [.. Sections, section] };
    public Document WithStyles(params StyleDefinition[] styles) => this with { Styles = styles };
    public Document WithNumbering(params NumberingDefinition[] definitions) => this with { Numbering = definitions };
    public Document WithProtection(DocumentProtection protection) => this with { Protection = protection };

    /// <summary>Attaches a VBA project, defensively copying <paramref
    /// name="vbaProjectBytes"/> so later mutations to the caller's array don't leak into
    /// this document.</summary>
    public Document WithVbaProject(byte[] vbaProjectBytes) => this with { VbaProject = (byte[])vbaProjectBytes.Clone() };

    public Document WithDocumentProperties(DocumentProperties properties) => this with { Properties = properties };
    public Document WithTableStyles(params TableStyleDefinition[] definitions) => this with { TableStyles = definitions };
}
