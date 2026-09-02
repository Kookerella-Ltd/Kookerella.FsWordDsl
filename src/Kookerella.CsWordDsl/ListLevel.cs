namespace Kookerella.CsWordDsl;

/// <summary>
/// One level of a (potentially multi-level) list definition. <see cref="Text"/> is the
/// level's raw pattern - a literal glyph for <see cref="NumberFormatKind.Bullet"/>, or a
/// <c>"%1."</c>-style pattern for the numbered formats (<c>%1</c> is replaced with this
/// level's own counter; a deeper level's pattern can reference an ancestor level's counter
/// too, e.g. <c>"%1.%2"</c> - this wrapper does not validate that the pattern's
/// placeholders match the level nesting).
/// </summary>
public sealed record ListLevel
{
    public required NumberFormatKind Format { get; init; }
    public required string Text { get; init; }

    /// <summary>Points.</summary>
    public double? IndentLeft { get; init; }

    /// <summary>Points.</summary>
    public double? HangingIndent { get; init; }

    public int? StartAt { get; init; }
}

/// <summary><see cref="Id"/> is the number a <see cref="Paragraph"/>'s numbering
/// references, scoped to the <see cref="Document"/> it lives on.</summary>
public sealed record NumberingDefinition(int Id, IReadOnlyList<ListLevel> Levels)
{
    /// <summary>A single-level bullet list definition using Word's own conventional bullet
    /// glyph.</summary>
    public static NumberingDefinition BulletList(int id) =>
        new(id, [new ListLevel { Format = new NumberFormatKind.Bullet((char)0xF0B7, "Symbol"), Text = ((char)0xF0B7).ToString(), IndentLeft = 36.0, HangingIndent = 18.0 }]);

    /// <summary>A single-level decimal-numbered list definition ("1.", "2.", "3.", ...).
    /// </summary>
    public static NumberingDefinition NumberedList(int id) =>
        new(id, [new ListLevel { Format = new NumberFormatKind.Decimal(), Text = "%1.", IndentLeft = 36.0, HangingIndent = 18.0, StartAt = 1 }]);

    /// <summary>A multi-level decimal-numbered outline list ("1.", "1.1.", "1.1.1.", ...).
    /// <paramref name="levelCount"/> must be between 1 and 9.</summary>
    public static NumberingDefinition MultiLevelNumberedList(int id, int levelCount)
    {
        var levels = new List<ListLevel>();

        for (var i = 1; i <= levelCount; i++)
        {
            var text = string.Join(".", Enumerable.Range(1, i).Select(n => $"%{n}")) + ".";
            levels.Add(new ListLevel { Format = new NumberFormatKind.Decimal(), Text = text, IndentLeft = 36.0 * i, HangingIndent = 18.0, StartAt = 1 });
        }

        return new NumberingDefinition(id, levels);
    }
}
