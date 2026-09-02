using Xunit;
using static Kookerella.CsWordDsl.Tests.TestHelpers;

namespace Kookerella.CsWordDsl.Tests;

/// <summary>
/// Reloads every scenario the F# test suite's own <c>verifyScenarioNamed</c> harness
/// already built and checked in under <c>Examples/*/output.docx</c> (see this project's
/// .csproj for how those files are linked in rather than copy-pasted) - broad coverage of
/// this wrapper's read path across every feature the F# core models, without re-authoring
/// each scenario a second time. Each one just needs to load without throwing and produce a
/// non-empty document; <c>DocumentTests</c> covers the specific-value assertions per
/// feature.
/// </summary>
public class ExampleTests
{
    public static IEnumerable<object[]> DocxScenarios =>
        new[]
        {
            "BasicParagraphsAndRuns", "Bookmark", "Bookmark_MultiParagraph", "BulletList",
            "Comments", "Comments_MultiParagraph", "ContentControls", "DocumentProperties",
            "FootnotesAndEndnotes", "HeaderFooterDefault", "HeaderFooterFirstPageDifferent",
            "Hyperlink_External", "Hyperlink_Internal", "Image", "MultiLevelNumberedList",
            "MultipleSections", "NamedStyles", "NumberedList", "PageSetupLandscape",
            "ParagraphFormatting", "Table_Basic", "Table_BordersAndStyle",
            "Table_CustomStyleAndHeaderRow", "Table_MergedCells", "TrackedChanges"
        }.Select(name => new object[] { name });

    [Theory]
    [MemberData(nameof(DocxScenarios))]
    public void Loads_the_fsharp_suites_own_example(string scenarioName)
    {
        var path = ExamplePath(scenarioName, "output.docx");
        var doc = DocumentIO.Load(path);

        Assert.NotEmpty(doc.Sections);
        Assert.NotEmpty(doc.Sections.SelectMany(s => s.Body));
    }

    public static IEnumerable<object[]> DocmScenarios =>
        new[] { "DocumentProtectionReadOnly", "Macro" }.Select(name => new object[] { name });

    [Theory]
    [MemberData(nameof(DocmScenarios))]
    public void Loads_the_fsharp_suites_own_macro_enabled_example(string scenarioName)
    {
        var path = ExamplePath(scenarioName, "output.docm");
        var doc = DocumentIO.Load(path);

        Assert.NotEmpty(doc.Sections);
    }

    [Fact]
    public void Loaded_content_control_example_has_all_five_control_kinds()
    {
        var doc = DocumentIO.Load(ExamplePath("ContentControls", "output.docx"));
        var allBlocks = doc.Sections.SelectMany(s => s.Body).SelectMany(FlattenBlocks).ToList();

        var inlineControls = allBlocks
            .OfType<Block.ParagraphBlock>()
            .SelectMany(p => p.Para.Inlines)
            .OfType<Inline.ContentControl>()
            .Select(cc => cc.Props.Type)
            .ToList();

        Assert.Contains(inlineControls, t => t is ContentControlType.PlainText);
        Assert.Contains(inlineControls, t => t is ContentControlType.DropDown);
        Assert.Contains(inlineControls, t => t is ContentControlType.Date);
        Assert.Contains(inlineControls, t => t is ContentControlType.CheckBox);

        // The scenario's rich-text example is block-level (ContentControlBlock), not
        // inline - a separate F# DU case from the four checked above.
        Assert.Contains(allBlocks, b => b is Block.ContentControlBlock cc && cc.Props.Type is ContentControlType.RichText);
    }

    private static IEnumerable<Block> FlattenBlocks(Block block)
    {
        yield return block;

        switch (block)
        {
            case Block.TableBlock t:
                foreach (var b in t.Entry.Rows.SelectMany(r => r.Cells).SelectMany(c => c.Content).SelectMany(FlattenBlocks))
                    yield return b;
                break;
            case Block.ContentControlBlock cc:
                foreach (var b in cc.Content.SelectMany(FlattenBlocks))
                    yield return b;
                break;
        }
    }
}
