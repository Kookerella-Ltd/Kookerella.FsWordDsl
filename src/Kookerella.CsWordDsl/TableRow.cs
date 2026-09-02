namespace Kookerella.CsWordDsl;

public sealed record TableRow
{
    public required IReadOnlyList<TableCell> Cells { get; init; }

    /// <summary>Points.</summary>
    public double? Height { get; init; }

    /// <summary>Repeats this row at the top of every page the table spans - meaningful
    /// only on a table's leading row(s).</summary>
    public bool RepeatAsHeader { get; init; }

    public static TableRow Of(IReadOnlyList<TableCell> cells, double? height = null, bool repeatAsHeader = false) =>
        new() { Cells = cells, Height = height, RepeatAsHeader = repeatAsHeader };
}

/// <summary><see cref="ColumnWidths"/> gives the table's own grid - one entry per column,
/// in points; a row's own cells should sum to the same column count accounting for any
/// <see cref="TableCellProps.GridSpan"/>.</summary>
public sealed record TableEntry
{
    public required IReadOnlyList<TableRow> Rows { get; init; }
    public required IReadOnlyList<double> ColumnWidths { get; init; }
    public TableStyleRef? Style { get; init; }
    public TableBorders? Borders { get; init; }

    /// <summary>The table's own default cell margins - see <see cref="CellMargins"/>'s own
    /// doc comment for what this does and doesn't cover.</summary>
    public CellMargins? CellMargins { get; init; }
}
