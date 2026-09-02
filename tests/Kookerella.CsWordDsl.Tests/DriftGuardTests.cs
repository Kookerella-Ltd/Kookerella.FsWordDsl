using Microsoft.FSharp.Reflection;
using Xunit;
using Fs = Kookerella.FsWordDsl;

namespace Kookerella.CsWordDsl.Tests;

/// <summary>
/// A structural tripwire against the specific failure mode this project's F# core has hit
/// repeatedly during its own development: a case gets added to an F# discriminated union,
/// and this C# wrapper silently doesn't grow a matching case, because nothing forces
/// anyone to notice. This doesn't verify the C# side is *correct* - only that its case
/// count hasn't fallen behind the F# core's. Mirrors the identically-named test class in
/// the sibling Kookerella.CsOpenXmlDsl.Tests project.
///
/// Deliberately count-based rather than name-based: several mirrors rename cases on
/// purpose (e.g. F#'s <c>ContentControlType.PlainTextControl</c> becomes plain <c>
/// ContentControlType.PlainText</c> here, since "Control" is already implied by the
/// containing type), so comparing exact case names would false-positive on those. A count
/// mismatch is a strictly weaker but far more robust signal - it still catches an
/// added-and-forgotten case, it just doesn't name which one.
/// </summary>
public class DriftGuardTests
{
    /// <summary>F# cases with no C# counterpart on purpose. Adding an entry here without a
    /// corresponding doc trail (a MAPPING.md gap, or a wrapper doc comment explaining the
    /// omission) defeats the point of this test.</summary>
    private static readonly Dictionary<Type, string[]> KnownGaps = new();

    private static int FsCaseCount(Type fsUnionType)
    {
        var gapCount = KnownGaps.TryGetValue(fsUnionType, out var gaps) ? gaps.Length : 0;
        return FSharpType.GetUnionCases(fsUnionType, null).Length - gapCount;
    }

    private static int CsEnumCaseCount(Type csEnumType) => Enum.GetNames(csEnumType).Length;

    /// <summary>Counts a closed hierarchy's case types the same way this wrapper builds
    /// them (nested sealed records directly under an abstract base).</summary>
    private static int CsClosedHierarchyCaseCount(Type csAbstractBaseType) =>
        csAbstractBaseType.Assembly.GetTypes().Count(t => t.BaseType == csAbstractBaseType);

    public static IEnumerable<object[]> EnumMirrors =>
        new List<(Type Fs, Type Cs)>
        {
            (typeof(Fs.PageSetup.PageOrientation), typeof(PageOrientation)),
            (typeof(Fs.PageSetup.SectionBreakType), typeof(SectionBreakType)),
            (typeof(Fs.Protection.EditRestriction), typeof(EditRestriction)),
            (typeof(Fs.Revisions.RevisionKind), typeof(RevisionKind)),
            (typeof(Fs.ContentControls.ContentControlLock), typeof(ContentControlLock)),
            (typeof(Fs.Tables.VerticalMergeKind), typeof(VerticalMergeKind)),
            (typeof(Fs.NamedStyles.StyleTargetType), typeof(StyleTargetType)),
            (typeof(Fs.PageSetup.NoteNumberRestart), typeof(NoteNumberRestart)),
            (typeof(Fs.Images.ImageFormat), typeof(ImageFormat)),
            (typeof(Fs.Styles.ThemeColorKind), typeof(ThemeColorKind)),
            (typeof(Fs.Styles.HighlightColor), typeof(HighlightColor)),
            (typeof(Fs.Styles.VerticalPosition), typeof(VerticalPosition)),
            (typeof(Fs.Styles.TabLeader), typeof(TabLeader)),
            (typeof(Fs.Styles.ParagraphAlignment), typeof(ParagraphAlignment)),
        }.Select(p => new object[] { p.Fs, p.Cs });

    [Theory]
    [MemberData(nameof(EnumMirrors))]
    public void Enum_case_count_matches_fs_core(Type fsType, Type csType) =>
        Assert.Equal(FsCaseCount(fsType), CsEnumCaseCount(csType));

    public static IEnumerable<object[]> ClosedHierarchyMirrors =>
        new List<(Type Fs, Type Cs)>
        {
            (typeof(Fs.Model.Inline), typeof(Inline)),
            (typeof(Fs.Model.Block), typeof(Block)),
            (typeof(Fs.Styles.Color), typeof(Color)),
            (typeof(Fs.Styles.UnderlineStyle), typeof(UnderlineStyle)),
            (typeof(Fs.Styles.BorderLineStyle), typeof(BorderLineStyle)),
            (typeof(Fs.Styles.TabStopAlignment), typeof(TabStopAlignment)),
            (typeof(Fs.Styles.LineSpacingRule), typeof(LineSpacingRule)),
            (typeof(Fs.Numbering.NumberFormatKind), typeof(NumberFormatKind)),
            (typeof(Fs.Hyperlinks.HyperlinkTarget), typeof(HyperlinkTarget)),
            (typeof(Fs.PageSetup.PageSize), typeof(PageSize)),
            (typeof(Fs.ContentControls.ContentControlType), typeof(ContentControlType)),
        }.Select(p => new object[] { p.Fs, p.Cs });

    [Theory]
    [MemberData(nameof(ClosedHierarchyMirrors))]
    public void Closed_hierarchy_case_count_matches_fs_core(Type fsType, Type csType) =>
        Assert.Equal(FsCaseCount(fsType), CsClosedHierarchyCaseCount(csType));
}
