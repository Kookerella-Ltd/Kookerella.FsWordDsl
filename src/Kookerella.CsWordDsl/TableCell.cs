namespace Kookerella.CsWordDsl;

/// <summary>One table cell. <see cref="Content"/> is almost always a single paragraph in
/// practice (Word requires at least one paragraph per cell even when empty), but nothing
/// here enforces that.</summary>
public sealed record TableCell
{
    public required IReadOnlyList<Block> Content { get; init; }
    public TableCellProps Props { get; init; } = TableCellProps.Default;

    public static TableCell Of(IReadOnlyList<Block> content, TableCellProps? props = null) =>
        new() { Content = content, Props = props ?? TableCellProps.Default };

    public TableCell WithProps(TableCellProps props) => this with { Props = props };
}
