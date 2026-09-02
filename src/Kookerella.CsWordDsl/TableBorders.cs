namespace Kookerella.CsWordDsl;

/// <summary>Reuses <see cref="BorderStyle"/>'s four sides and adds the two "inside
/// gridline" sides a multi-row/column table has that a single bordered box doesn't.</summary>
public sealed record TableBorders
{
    public BorderStyle Outer { get; init; } = BorderStyle.None;
    public BorderSide? InsideHorizontal { get; init; }
    public BorderSide? InsideVertical { get; init; }

    public static readonly TableBorders None = new();

    public TableBorders WithOuter(BorderStyle outer) => this with { Outer = outer };
    public TableBorders WithInsideHorizontal(BorderSide side) => this with { InsideHorizontal = side };
    public TableBorders WithInsideVertical(BorderSide side) => this with { InsideVertical = side };
}

/// <summary>A cell's vertical merge state - <see cref="Restart"/> begins a new merged
/// group (this cell is the visible one, spanning downward), <see cref="Continue"/> merges
/// into the group started by the nearest <see cref="Restart"/> above it in the same
/// column.</summary>
public enum VerticalMergeKind
{
    Restart,
    Continue
}

/// <summary>Cell margins - the table-wide default and a single cell's own override share
/// this exact shape. All fields in points.</summary>
public sealed record CellMargins
{
    public double? Top { get; init; }
    public double? Bottom { get; init; }
    public double? Left { get; init; }
    public double? Right { get; init; }

    public static readonly CellMargins Default = new();
}

/// <summary>Per-cell overrides. <see cref="GridSpan"/> (horizontal merge) and <see
/// cref="VerticalMerge"/> (vertical merge) are independent and can combine on the same
/// cell. <see langword="null"/> on <see cref="Shading"/>/<see cref="Borders"/>/<see
/// cref="Width"/>/<see cref="Margins"/> means "inherit from the table."</summary>
public sealed record TableCellProps
{
    public int? GridSpan { get; init; }
    public VerticalMergeKind? VerticalMerge { get; init; }
    public Color? Shading { get; init; }
    public TableBorders? Borders { get; init; }

    /// <summary>Points.</summary>
    public double? Width { get; init; }

    public CellMargins? Margins { get; init; }

    public static readonly TableCellProps Default = new();

    public TableCellProps WithGridSpan(int columns) => this with { GridSpan = columns };
    public TableCellProps WithVerticalMerge(VerticalMergeKind merge) => this with { VerticalMerge = merge };
    public TableCellProps WithShading(Color color) => this with { Shading = color };
    public TableCellProps WithBorders(TableBorders borders) => this with { Borders = borders };
    public TableCellProps WithWidth(double points) => this with { Width = points };
    public TableCellProps WithMargins(CellMargins margins) => this with { Margins = margins };
}

/// <summary>A reference to a table style by name - either a built-in like <c>"TableGrid"</c>,
/// or a custom <see cref="TableStyleDefinition"/>.</summary>
public sealed record TableStyleRef
{
    public required string Name { get; init; }
    public bool FirstRowBanding { get; init; }
    public bool LastRowBanding { get; init; }
    public bool BandedRows { get; init; }
    public bool BandedColumns { get; init; }

    public static TableStyleRef Named(string name) => new() { Name = name };
}
