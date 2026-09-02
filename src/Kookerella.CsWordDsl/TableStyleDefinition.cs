namespace Kookerella.CsWordDsl;

/// <summary>One conditional-formatting region within a custom table style. This wrapper
/// models all thirteen of OOXML's regions.</summary>
public sealed record TableStyleRegion
{
    public RunStyle? RunFormat { get; init; }
    public ParagraphFormat? ParaFormat { get; init; }

    /// <summary>This region's own cell background - independent of <see
    /// cref="RunFormat"/>/<see cref="ParaFormat"/>, which cover text/paragraph appearance
    /// only.</summary>
    public Color? CellShading { get; init; }

    public static readonly TableStyleRegion None = new();

    public TableStyleRegion WithRunFormat(RunStyle format) => this with { RunFormat = format };
    public TableStyleRegion WithParaFormat(ParagraphFormat format) => this with { ParaFormat = format };
    public TableStyleRegion WithCellShading(Color color) => this with { CellShading = color };
}

/// <summary>
/// A custom table style definition - unlike <see cref="TableStyleRef"/>, which only
/// references a style by name, this actually defines one. Add it to <see
/// cref="Document.TableStyles"/> and reference its <see cref="Id"/> from a <see
/// cref="TableStyleRef.Name"/> the same way you'd reference a built-in name like
/// <c>"TableGrid"</c>.
/// </summary>
public sealed record TableStyleDefinition
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public string? BasedOn { get; init; }

    /// <summary>The style's own base table borders.</summary>
    public TableBorders? Borders { get; init; }

    public TableStyleRegion WholeTable { get; init; } = TableStyleRegion.None;
    public TableStyleRegion FirstRow { get; init; } = TableStyleRegion.None;
    public TableStyleRegion LastRow { get; init; } = TableStyleRegion.None;
    public TableStyleRegion FirstColumn { get; init; } = TableStyleRegion.None;
    public TableStyleRegion LastColumn { get; init; } = TableStyleRegion.None;

    /// <summary>The odd/first band of each banding axis.</summary>
    public TableStyleRegion BandedRow { get; init; } = TableStyleRegion.None;
    public TableStyleRegion BandedColumn { get; init; } = TableStyleRegion.None;

    /// <summary>The even/second band of each banding axis.</summary>
    public TableStyleRegion BandedRow2 { get; init; } = TableStyleRegion.None;
    public TableStyleRegion BandedColumn2 { get; init; } = TableStyleRegion.None;

    /// <summary>The four corner cells - where a row-band region and a column-band region
    /// would otherwise overlap.</summary>
    public TableStyleRegion NorthEastCell { get; init; } = TableStyleRegion.None;
    public TableStyleRegion NorthWestCell { get; init; } = TableStyleRegion.None;
    public TableStyleRegion SouthEastCell { get; init; } = TableStyleRegion.None;
    public TableStyleRegion SouthWestCell { get; init; } = TableStyleRegion.None;

    public TableStyleDefinition WithBasedOn(string basedOnId) => this with { BasedOn = basedOnId };
    public TableStyleDefinition WithBorders(TableBorders borders) => this with { Borders = borders };
    public TableStyleDefinition WithWholeTable(TableStyleRegion region) => this with { WholeTable = region };
    public TableStyleDefinition WithFirstRow(TableStyleRegion region) => this with { FirstRow = region };
    public TableStyleDefinition WithLastRow(TableStyleRegion region) => this with { LastRow = region };
    public TableStyleDefinition WithFirstColumn(TableStyleRegion region) => this with { FirstColumn = region };
    public TableStyleDefinition WithLastColumn(TableStyleRegion region) => this with { LastColumn = region };
    public TableStyleDefinition WithBandedRow(TableStyleRegion region) => this with { BandedRow = region };
    public TableStyleDefinition WithBandedColumn(TableStyleRegion region) => this with { BandedColumn = region };
    public TableStyleDefinition WithBandedRow2(TableStyleRegion region) => this with { BandedRow2 = region };
    public TableStyleDefinition WithBandedColumn2(TableStyleRegion region) => this with { BandedColumn2 = region };
    public TableStyleDefinition WithNorthEastCell(TableStyleRegion region) => this with { NorthEastCell = region };
    public TableStyleDefinition WithNorthWestCell(TableStyleRegion region) => this with { NorthWestCell = region };
    public TableStyleDefinition WithSouthEastCell(TableStyleRegion region) => this with { SouthEastCell = region };
    public TableStyleDefinition WithSouthWestCell(TableStyleRegion region) => this with { SouthWestCell = region };
}
