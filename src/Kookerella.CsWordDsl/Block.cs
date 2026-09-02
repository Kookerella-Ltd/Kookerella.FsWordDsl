namespace Kookerella.CsWordDsl;

/// <summary>
/// A block-level unit of document content - a paragraph, a table, or a content control
/// wrapping either. Mirrors the F# core's own <c>Block</c> discriminated union.
/// </summary>
public abstract record Block
{
    private Block() { }

    public sealed record ParagraphBlock(Paragraph Para) : Block;

    public sealed record TableBlock(TableEntry Entry) : Block;

    /// <summary>The block-level counterpart to <see cref="Inline.ContentControl"/> -
    /// wrapping whole paragraphs/tables rather than sitting inside one paragraph.</summary>
    public sealed record ContentControlBlock(ContentControlProps Props, IReadOnlyList<Block> Content) : Block;

    /// <summary>Builds a <see cref="Paragraph"/> and wraps it in one call, matching the F#
    /// core's own <c>DocumentDsl.para</c> convenience.</summary>
    public static Block Paragraph(IReadOnlyList<Inline> inlines, string? styleId = null, ParagraphFormat? format = null, (int NumId, int Level)? numbering = null, Revision? markRevision = null) =>
        new ParagraphBlock(new global::Kookerella.CsWordDsl.Paragraph { Inlines = inlines, StyleId = styleId, Format = format, Numbering = numbering, MarkRevision = markRevision });

    /// <summary>Builds a <see cref="TableEntry"/> and wraps it in one call, matching the
    /// F# core's own <c>DocumentDsl.table</c> convenience.</summary>
    public static Block Table(IReadOnlyList<TableRow> rows, IReadOnlyList<double> columnWidths, TableStyleRef? style = null, TableBorders? borders = null, CellMargins? cellMargins = null) =>
        new TableBlock(new TableEntry { Rows = rows, ColumnWidths = columnWidths, Style = style, Borders = borders, CellMargins = cellMargins });
}
