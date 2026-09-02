namespace Kookerella.CsWordDsl;

/// <summary><see cref="Default"/>/<see cref="First"/>/<see cref="Even"/> mirror Word's own
/// three header/footer variants exactly - <see cref="First"/> shows only on a section's
/// first page, <see cref="Even"/> only on even pages. <see cref="Default"/> covers odd
/// pages when <see cref="Even"/> is set, or every page otherwise.</summary>
public sealed record HeaderFooterSet
{
    public IReadOnlyList<Block>? Default { get; init; }
    public IReadOnlyList<Block>? First { get; init; }
    public IReadOnlyList<Block>? Even { get; init; }

    public static readonly HeaderFooterSet None = new();

    public HeaderFooterSet WithDefault(IReadOnlyList<Block> blocks) => this with { Default = blocks };
    public HeaderFooterSet WithFirst(IReadOnlyList<Block> blocks) => this with { First = blocks };
    public HeaderFooterSet WithEven(IReadOnlyList<Block> blocks) => this with { Even = blocks };
}
