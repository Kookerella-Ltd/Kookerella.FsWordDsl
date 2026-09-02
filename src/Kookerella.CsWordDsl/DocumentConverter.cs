using Microsoft.FSharp.Collections;
using Microsoft.FSharp.Core;
using Fs = Kookerella.FsWordDsl;

namespace Kookerella.CsWordDsl;

/// <summary>
/// Pure translation between this wrapper's immutable C# records and the F# core's own
/// <c>Document</c>/<c>Section</c>/<c>Block</c>/<c>Inline</c>/... values - no I/O happens
/// here at all (see <see cref="DocumentIO"/> for the one place it does). Internal: callers
/// only ever see the C# <see cref="Document"/> shape, never the F# types underneath. F#
/// types are referenced via the <c>Fs</c> alias throughout rather than a blanket
/// <c>using</c>, since e.g. <c>Fs.Model.Document</c>/<c>Fs.Model.Paragraph</c> would
/// otherwise collide with this assembly's own same-named types.
/// <para>
/// A note on F#-compiled shapes, confirmed by reflection rather than assumed (see
/// <c>CLAUDE.md</c>'s own warning about this SDK/interop combination): an F# record's
/// fields become PascalCase properties reachable through a normal positional constructor;
/// a discriminated union case with data becomes a nested type constructed via a
/// <c>New&lt;CaseName&gt;</c> static factory (a case with no data is a static property
/// instead, not a method), and that nested case type's own fields keep whatever casing
/// their field label had in the F# source (lowercase, since every case here names its
/// fields lowercase) - a genuinely different capitalization convention from plain records,
/// easy to get wrong without checking.
/// </para>
/// </summary>
internal static class DocumentConverter
{
    // ----- FSharpOption / FSharpList / tuple helpers ---------------------------------------

    private static FSharpOption<T> ToOpt<T>(T? value) where T : class =>
        value is null ? FSharpOption<T>.None : FSharpOption<T>.Some(value);

    private static FSharpOption<T> ToOptStruct<T>(T? value) where T : struct =>
        value.HasValue ? FSharpOption<T>.Some(value.Value) : FSharpOption<T>.None;

    private static T? FromOpt<T>(FSharpOption<T> option) where T : class =>
        option is not null && FSharpOption<T>.get_IsSome(option) ? option.Value : null;

    private static T? FromOptStruct<T>(FSharpOption<T> option) where T : struct =>
        option is not null && FSharpOption<T>.get_IsSome(option) ? option.Value : null;

    private static FSharpList<TFs> ToFsList<TCs, TFs>(IEnumerable<TCs>? items, Func<TCs, TFs> convert) =>
        ListModule.OfSeq(items is null ? [] : items.Select(convert));

    private static IReadOnlyList<TCs> FromFsList<TFs, TCs>(FSharpList<TFs> list, Func<TFs, TCs> convert) =>
        list.Select(convert).ToList();

    private static Tuple<byte, byte, byte> ToFsRgbTuple((byte R, byte G, byte B) t) => Tuple.Create(t.R, t.G, t.B);
    private static (byte, byte, byte) FromFsRgbTuple(Tuple<byte, byte, byte> t) => (t.Item1, t.Item2, t.Item3);
    private static Tuple<string, string> ToFsPair((string, string) t) => Tuple.Create(t.Item1, t.Item2);
    private static (string, string) FromFsPair(Tuple<string, string> t) => (t.Item1, t.Item2);

    // ----- Simple (no-data) discriminated unions --------------------------------------------

    private static Fs.PageSetup.PageOrientation ToFsPageOrientation(PageOrientation o) => o switch
    {
        PageOrientation.Landscape => Fs.PageSetup.PageOrientation.Landscape,
        _ => Fs.PageSetup.PageOrientation.Portrait
    };

    private static PageOrientation FromFsPageOrientation(Fs.PageSetup.PageOrientation o) =>
        o.IsLandscape ? PageOrientation.Landscape : PageOrientation.Portrait;

    private static Fs.PageSetup.SectionBreakType ToFsSectionBreakType(SectionBreakType t) => t switch
    {
        SectionBreakType.Continuous => Fs.PageSetup.SectionBreakType.ContinuousBreak,
        SectionBreakType.EvenPage => Fs.PageSetup.SectionBreakType.EvenPageBreak,
        SectionBreakType.OddPage => Fs.PageSetup.SectionBreakType.OddPageBreak,
        _ => Fs.PageSetup.SectionBreakType.NextPageBreak
    };

    private static SectionBreakType FromFsSectionBreakType(Fs.PageSetup.SectionBreakType t) =>
        t.IsContinuousBreak ? SectionBreakType.Continuous :
        t.IsEvenPageBreak ? SectionBreakType.EvenPage :
        t.IsOddPageBreak ? SectionBreakType.OddPage :
        SectionBreakType.NextPage;

    private static Fs.Protection.EditRestriction ToFsEditRestriction(EditRestriction e) => e switch
    {
        EditRestriction.CommentsOnly => Fs.Protection.EditRestriction.CommentsOnlyRestriction,
        EditRestriction.TrackedChangesOnly => Fs.Protection.EditRestriction.TrackedChangesOnlyRestriction,
        EditRestriction.FormsOnly => Fs.Protection.EditRestriction.FormsOnlyRestriction,
        _ => Fs.Protection.EditRestriction.ReadOnlyRestriction
    };

    private static EditRestriction FromFsEditRestriction(Fs.Protection.EditRestriction e) =>
        e.IsCommentsOnlyRestriction ? EditRestriction.CommentsOnly :
        e.IsTrackedChangesOnlyRestriction ? EditRestriction.TrackedChangesOnly :
        e.IsFormsOnlyRestriction ? EditRestriction.FormsOnly :
        EditRestriction.ReadOnly;

    private static Fs.Revisions.RevisionKind ToFsRevisionKind(RevisionKind k) =>
        k == RevisionKind.Deleted ? Fs.Revisions.RevisionKind.Deleted : Fs.Revisions.RevisionKind.Inserted;

    private static RevisionKind FromFsRevisionKind(Fs.Revisions.RevisionKind k) =>
        k.IsDeleted ? RevisionKind.Deleted : RevisionKind.Inserted;

    private static Fs.ContentControls.ContentControlLock ToFsContentControlLock(ContentControlLock l) => l switch
    {
        ContentControlLock.LockContentEditing => Fs.ContentControls.ContentControlLock.LockContentEditing,
        ContentControlLock.LockDeletionAndContentEditing => Fs.ContentControls.ContentControlLock.LockDeletionAndContentEditing,
        _ => Fs.ContentControls.ContentControlLock.LockDeletion
    };

    private static ContentControlLock FromFsContentControlLock(Fs.ContentControls.ContentControlLock l) =>
        l.IsLockContentEditing ? ContentControlLock.LockContentEditing :
        l.IsLockDeletionAndContentEditing ? ContentControlLock.LockDeletionAndContentEditing :
        ContentControlLock.LockDeletion;

    private static Fs.Tables.VerticalMergeKind ToFsVerticalMergeKind(VerticalMergeKind k) =>
        k == VerticalMergeKind.Continue ? Fs.Tables.VerticalMergeKind.ContinueMerge : Fs.Tables.VerticalMergeKind.RestartMerge;

    private static VerticalMergeKind FromFsVerticalMergeKind(Fs.Tables.VerticalMergeKind k) =>
        k.IsContinueMerge ? VerticalMergeKind.Continue : VerticalMergeKind.Restart;

    private static Fs.NamedStyles.StyleTargetType ToFsStyleTargetType(StyleTargetType t) =>
        t == StyleTargetType.Character ? Fs.NamedStyles.StyleTargetType.CharacterStyleType : Fs.NamedStyles.StyleTargetType.ParagraphStyleType;

    private static StyleTargetType FromFsStyleTargetType(Fs.NamedStyles.StyleTargetType t) =>
        t.IsCharacterStyleType ? StyleTargetType.Character : StyleTargetType.Paragraph;

    private static Fs.PageSetup.NoteNumberRestart ToFsNoteNumberRestart(NoteNumberRestart r) => r switch
    {
        NoteNumberRestart.EachSection => Fs.PageSetup.NoteNumberRestart.RestartEachSection,
        NoteNumberRestart.EachPage => Fs.PageSetup.NoteNumberRestart.RestartEachPage,
        _ => Fs.PageSetup.NoteNumberRestart.ContinuousRestart
    };

    private static NoteNumberRestart FromFsNoteNumberRestart(Fs.PageSetup.NoteNumberRestart r) =>
        r.IsRestartEachSection ? NoteNumberRestart.EachSection :
        r.IsRestartEachPage ? NoteNumberRestart.EachPage :
        NoteNumberRestart.Continuous;

    private static Fs.Images.ImageFormat ToFsImageFormat(ImageFormat f) => f switch
    {
        ImageFormat.Jpeg => Fs.Images.ImageFormat.Jpeg,
        ImageFormat.Gif => Fs.Images.ImageFormat.Gif,
        ImageFormat.Bmp => Fs.Images.ImageFormat.Bmp,
        _ => Fs.Images.ImageFormat.Png
    };

    private static ImageFormat FromFsImageFormat(Fs.Images.ImageFormat f) =>
        f.IsJpeg ? ImageFormat.Jpeg :
        f.IsGif ? ImageFormat.Gif :
        f.IsBmp ? ImageFormat.Bmp :
        ImageFormat.Png;

    private static Fs.Styles.ThemeColorKind ToFsThemeColorKind(ThemeColorKind k) => k switch
    {
        ThemeColorKind.Dark1 => Fs.Styles.ThemeColorKind.Dark1Theme,
        ThemeColorKind.Light1 => Fs.Styles.ThemeColorKind.Light1Theme,
        ThemeColorKind.Dark2 => Fs.Styles.ThemeColorKind.Dark2Theme,
        ThemeColorKind.Light2 => Fs.Styles.ThemeColorKind.Light2Theme,
        ThemeColorKind.Accent1 => Fs.Styles.ThemeColorKind.Accent1Theme,
        ThemeColorKind.Accent2 => Fs.Styles.ThemeColorKind.Accent2Theme,
        ThemeColorKind.Accent3 => Fs.Styles.ThemeColorKind.Accent3Theme,
        ThemeColorKind.Accent4 => Fs.Styles.ThemeColorKind.Accent4Theme,
        ThemeColorKind.Accent5 => Fs.Styles.ThemeColorKind.Accent5Theme,
        ThemeColorKind.Accent6 => Fs.Styles.ThemeColorKind.Accent6Theme,
        ThemeColorKind.Hyperlink => Fs.Styles.ThemeColorKind.HyperlinkTheme,
        ThemeColorKind.FollowedHyperlink => Fs.Styles.ThemeColorKind.FollowedHyperlinkTheme,
        ThemeColorKind.Background1 => Fs.Styles.ThemeColorKind.Background1Theme,
        ThemeColorKind.Text1 => Fs.Styles.ThemeColorKind.Text1Theme,
        ThemeColorKind.Background2 => Fs.Styles.ThemeColorKind.Background2Theme,
        _ => Fs.Styles.ThemeColorKind.Text2Theme
    };

    private static ThemeColorKind FromFsThemeColorKind(Fs.Styles.ThemeColorKind k) =>
        k.IsDark1Theme ? ThemeColorKind.Dark1 :
        k.IsLight1Theme ? ThemeColorKind.Light1 :
        k.IsDark2Theme ? ThemeColorKind.Dark2 :
        k.IsLight2Theme ? ThemeColorKind.Light2 :
        k.IsAccent1Theme ? ThemeColorKind.Accent1 :
        k.IsAccent2Theme ? ThemeColorKind.Accent2 :
        k.IsAccent3Theme ? ThemeColorKind.Accent3 :
        k.IsAccent4Theme ? ThemeColorKind.Accent4 :
        k.IsAccent5Theme ? ThemeColorKind.Accent5 :
        k.IsAccent6Theme ? ThemeColorKind.Accent6 :
        k.IsHyperlinkTheme ? ThemeColorKind.Hyperlink :
        k.IsFollowedHyperlinkTheme ? ThemeColorKind.FollowedHyperlink :
        k.IsBackground1Theme ? ThemeColorKind.Background1 :
        k.IsText1Theme ? ThemeColorKind.Text1 :
        k.IsBackground2Theme ? ThemeColorKind.Background2 :
        ThemeColorKind.Text2;

    private static Fs.Styles.HighlightColor ToFsHighlightColor(HighlightColor h) => h switch
    {
        HighlightColor.Green => Fs.Styles.HighlightColor.HlGreen,
        HighlightColor.Cyan => Fs.Styles.HighlightColor.HlCyan,
        HighlightColor.Magenta => Fs.Styles.HighlightColor.HlMagenta,
        HighlightColor.Blue => Fs.Styles.HighlightColor.HlBlue,
        HighlightColor.Red => Fs.Styles.HighlightColor.HlRed,
        HighlightColor.DarkBlue => Fs.Styles.HighlightColor.HlDarkBlue,
        HighlightColor.DarkCyan => Fs.Styles.HighlightColor.HlDarkCyan,
        HighlightColor.DarkGreen => Fs.Styles.HighlightColor.HlDarkGreen,
        HighlightColor.DarkMagenta => Fs.Styles.HighlightColor.HlDarkMagenta,
        HighlightColor.DarkRed => Fs.Styles.HighlightColor.HlDarkRed,
        HighlightColor.DarkYellow => Fs.Styles.HighlightColor.HlDarkYellow,
        HighlightColor.DarkGray => Fs.Styles.HighlightColor.HlDarkGray,
        HighlightColor.LightGray => Fs.Styles.HighlightColor.HlLightGray,
        HighlightColor.Black => Fs.Styles.HighlightColor.HlBlack,
        _ => Fs.Styles.HighlightColor.HlYellow
    };

    private static HighlightColor FromFsHighlightColor(Fs.Styles.HighlightColor h) =>
        h.IsHlGreen ? HighlightColor.Green :
        h.IsHlCyan ? HighlightColor.Cyan :
        h.IsHlMagenta ? HighlightColor.Magenta :
        h.IsHlBlue ? HighlightColor.Blue :
        h.IsHlRed ? HighlightColor.Red :
        h.IsHlDarkBlue ? HighlightColor.DarkBlue :
        h.IsHlDarkCyan ? HighlightColor.DarkCyan :
        h.IsHlDarkGreen ? HighlightColor.DarkGreen :
        h.IsHlDarkMagenta ? HighlightColor.DarkMagenta :
        h.IsHlDarkRed ? HighlightColor.DarkRed :
        h.IsHlDarkYellow ? HighlightColor.DarkYellow :
        h.IsHlDarkGray ? HighlightColor.DarkGray :
        h.IsHlLightGray ? HighlightColor.LightGray :
        h.IsHlBlack ? HighlightColor.Black :
        HighlightColor.Yellow;

    private static Fs.Styles.VerticalPosition ToFsVerticalPosition(VerticalPosition v) =>
        v == VerticalPosition.Subscript ? Fs.Styles.VerticalPosition.Subscript : Fs.Styles.VerticalPosition.Superscript;

    private static VerticalPosition FromFsVerticalPosition(Fs.Styles.VerticalPosition v) =>
        v.IsSubscript ? VerticalPosition.Subscript : VerticalPosition.Superscript;

    private static Fs.Styles.TabLeader ToFsTabLeader(TabLeader l) => l switch
    {
        TabLeader.Dot => Fs.Styles.TabLeader.DotLeader,
        TabLeader.Hyphen => Fs.Styles.TabLeader.HyphenLeader,
        TabLeader.Underscore => Fs.Styles.TabLeader.UnderscoreLeader,
        TabLeader.Heavy => Fs.Styles.TabLeader.HeavyLeader,
        TabLeader.MiddleDot => Fs.Styles.TabLeader.MiddleDotLeader,
        _ => Fs.Styles.TabLeader.NoLeader
    };

    private static TabLeader FromFsTabLeader(Fs.Styles.TabLeader l) =>
        l.IsDotLeader ? TabLeader.Dot :
        l.IsHyphenLeader ? TabLeader.Hyphen :
        l.IsUnderscoreLeader ? TabLeader.Underscore :
        l.IsHeavyLeader ? TabLeader.Heavy :
        l.IsMiddleDotLeader ? TabLeader.MiddleDot :
        TabLeader.None;

    private static Fs.Styles.ParagraphAlignment ToFsParagraphAlignment(ParagraphAlignment a) => a switch
    {
        ParagraphAlignment.Center => Fs.Styles.ParagraphAlignment.AlignCenter,
        ParagraphAlignment.Right => Fs.Styles.ParagraphAlignment.AlignRight,
        ParagraphAlignment.Justify => Fs.Styles.ParagraphAlignment.AlignJustify,
        _ => Fs.Styles.ParagraphAlignment.AlignLeft
    };

    private static ParagraphAlignment FromFsParagraphAlignment(Fs.Styles.ParagraphAlignment a) =>
        a.IsAlignCenter ? ParagraphAlignment.Center :
        a.IsAlignRight ? ParagraphAlignment.Right :
        a.IsAlignJustify ? ParagraphAlignment.Justify :
        ParagraphAlignment.Left;

    // ----- Value discriminated unions --------------------------------------------------------

    private static Fs.Styles.Color ToFsColor(Color c) => c switch
    {
        Color.Rgb rgb => Fs.Styles.Color.NewRgb(rgb.R, rgb.G, rgb.B),
        Color.Theme theme => Fs.Styles.Color.NewTheme(ToFsThemeColorKind(theme.Kind), ToFsRgbTuple((theme.Fallback.R, theme.Fallback.G, theme.Fallback.B)), ToOptStruct(theme.Tint), ToOptStruct(theme.Shade)),
        _ => Fs.Styles.Color.Auto
    };

    private static Color FromFsColor(Fs.Styles.Color c)
    {
        if (c.IsRgb)
        {
            var rgb = (Fs.Styles.Color.Rgb)c;
            return new Color.Rgb(rgb.red, rgb.green, rgb.blue);
        }

        if (c.IsTheme)
        {
            var theme = (Fs.Styles.Color.Theme)c;
            var (r, g, b) = FromFsRgbTuple(theme.fallback);
            return new Color.Theme(FromFsThemeColorKind(theme.kind), new Color.Rgb(r, g, b), FromOptStruct(theme.tint), FromOptStruct(theme.shade));
        }

        return new Color.Auto();
    }

    private static Fs.Styles.UnderlineStyle ToFsUnderlineStyle(UnderlineStyle u) => u switch
    {
        UnderlineStyle.Double => Fs.Styles.UnderlineStyle.DoubleUnderline,
        UnderlineStyle.Thick => Fs.Styles.UnderlineStyle.ThickUnderline,
        UnderlineStyle.Dotted => Fs.Styles.UnderlineStyle.DottedUnderline,
        UnderlineStyle.Dashed => Fs.Styles.UnderlineStyle.DashedUnderline,
        UnderlineStyle.Wavy => Fs.Styles.UnderlineStyle.WavyUnderline,
        UnderlineStyle.Other other => Fs.Styles.UnderlineStyle.NewOtherUnderline(other.Raw),
        _ => Fs.Styles.UnderlineStyle.SingleUnderline
    };

    private static UnderlineStyle FromFsUnderlineStyle(Fs.Styles.UnderlineStyle u) =>
        u.IsDoubleUnderline ? new UnderlineStyle.Double() :
        u.IsThickUnderline ? new UnderlineStyle.Thick() :
        u.IsDottedUnderline ? new UnderlineStyle.Dotted() :
        u.IsDashedUnderline ? new UnderlineStyle.Dashed() :
        u.IsWavyUnderline ? new UnderlineStyle.Wavy() :
        u.IsOtherUnderline ? new UnderlineStyle.Other(((Fs.Styles.UnderlineStyle.OtherUnderline)u).Item) :
        new UnderlineStyle.Single();

    private static Fs.Styles.BorderLineStyle ToFsBorderLineStyle(BorderLineStyle b) => b switch
    {
        BorderLineStyle.Thick => Fs.Styles.BorderLineStyle.ThickLine,
        BorderLineStyle.Double => Fs.Styles.BorderLineStyle.DoubleLine,
        BorderLineStyle.Dotted => Fs.Styles.BorderLineStyle.DottedLine,
        BorderLineStyle.Dashed => Fs.Styles.BorderLineStyle.DashedLine,
        BorderLineStyle.Wave => Fs.Styles.BorderLineStyle.WaveLine,
        BorderLineStyle.Other other => Fs.Styles.BorderLineStyle.NewOtherLine(other.Raw),
        _ => Fs.Styles.BorderLineStyle.SingleLine
    };

    private static BorderLineStyle FromFsBorderLineStyle(Fs.Styles.BorderLineStyle b) =>
        b.IsThickLine ? new BorderLineStyle.Thick() :
        b.IsDoubleLine ? new BorderLineStyle.Double() :
        b.IsDottedLine ? new BorderLineStyle.Dotted() :
        b.IsDashedLine ? new BorderLineStyle.Dashed() :
        b.IsWaveLine ? new BorderLineStyle.Wave() :
        b.IsOtherLine ? new BorderLineStyle.Other(((Fs.Styles.BorderLineStyle.OtherLine)b).Item) :
        new BorderLineStyle.Single();

    private static Fs.Styles.TabStopAlignment ToFsTabStopAlignment(TabStopAlignment a) => a switch
    {
        TabStopAlignment.Center => Fs.Styles.TabStopAlignment.CenterTab,
        TabStopAlignment.Right => Fs.Styles.TabStopAlignment.RightTab,
        TabStopAlignment.Decimal => Fs.Styles.TabStopAlignment.DecimalTab,
        TabStopAlignment.Bar => Fs.Styles.TabStopAlignment.BarTab,
        TabStopAlignment.Other other => Fs.Styles.TabStopAlignment.NewOtherTabAlignment(other.Raw),
        _ => Fs.Styles.TabStopAlignment.LeftTab
    };

    private static TabStopAlignment FromFsTabStopAlignment(Fs.Styles.TabStopAlignment a) =>
        a.IsCenterTab ? new TabStopAlignment.Center() :
        a.IsRightTab ? new TabStopAlignment.Right() :
        a.IsDecimalTab ? new TabStopAlignment.Decimal() :
        a.IsBarTab ? new TabStopAlignment.Bar() :
        a.IsOtherTabAlignment ? new TabStopAlignment.Other(((Fs.Styles.TabStopAlignment.OtherTabAlignment)a).Item) :
        new TabStopAlignment.Left();

    private static Fs.Styles.LineSpacingRule ToFsLineSpacingRule(LineSpacingRule r) => r switch
    {
        LineSpacingRule.OnePointFive => Fs.Styles.LineSpacingRule.OnePointFiveSpacing,
        LineSpacingRule.DoubleSpacing => Fs.Styles.LineSpacingRule.DoubleSpacing,
        LineSpacingRule.AtLeast atLeast => Fs.Styles.LineSpacingRule.NewAtLeastSpacing(atLeast.Points),
        LineSpacingRule.Exactly exactly => Fs.Styles.LineSpacingRule.NewExactlySpacing(exactly.Points),
        LineSpacingRule.Multiple multiple => Fs.Styles.LineSpacingRule.NewMultipleSpacing(multiple.Factor),
        _ => Fs.Styles.LineSpacingRule.SingleSpacing
    };

    private static LineSpacingRule FromFsLineSpacingRule(Fs.Styles.LineSpacingRule r) =>
        r.IsOnePointFiveSpacing ? new LineSpacingRule.OnePointFive() :
        r.IsDoubleSpacing ? new LineSpacingRule.DoubleSpacing() :
        r.IsAtLeastSpacing ? new LineSpacingRule.AtLeast(((Fs.Styles.LineSpacingRule.AtLeastSpacing)r).points) :
        r.IsExactlySpacing ? new LineSpacingRule.Exactly(((Fs.Styles.LineSpacingRule.ExactlySpacing)r).points) :
        r.IsMultipleSpacing ? new LineSpacingRule.Multiple(((Fs.Styles.LineSpacingRule.MultipleSpacing)r).factor) :
        new LineSpacingRule.Single();

    private static Fs.Numbering.NumberFormatKind ToFsNumberFormatKind(NumberFormatKind k) => k switch
    {
        NumberFormatKind.Bullet bullet => Fs.Numbering.NumberFormatKind.NewBulletFormat(bullet.Glyph, bullet.FontFamily),
        NumberFormatKind.LowerLetter => Fs.Numbering.NumberFormatKind.LowerLetterFormat,
        NumberFormatKind.UpperLetter => Fs.Numbering.NumberFormatKind.UpperLetterFormat,
        NumberFormatKind.LowerRoman => Fs.Numbering.NumberFormatKind.LowerRomanFormat,
        NumberFormatKind.UpperRoman => Fs.Numbering.NumberFormatKind.UpperRomanFormat,
        NumberFormatKind.Other other => Fs.Numbering.NumberFormatKind.NewOtherFormat(other.Raw),
        _ => Fs.Numbering.NumberFormatKind.DecimalFormat
    };

    private static NumberFormatKind FromFsNumberFormatKind(Fs.Numbering.NumberFormatKind k) =>
        k.IsBulletFormat ? new NumberFormatKind.Bullet(((Fs.Numbering.NumberFormatKind.BulletFormat)k).glyph, ((Fs.Numbering.NumberFormatKind.BulletFormat)k).fontFamily) :
        k.IsLowerLetterFormat ? new NumberFormatKind.LowerLetter() :
        k.IsUpperLetterFormat ? new NumberFormatKind.UpperLetter() :
        k.IsLowerRomanFormat ? new NumberFormatKind.LowerRoman() :
        k.IsUpperRomanFormat ? new NumberFormatKind.UpperRoman() :
        k.IsOtherFormat ? new NumberFormatKind.Other(((Fs.Numbering.NumberFormatKind.OtherFormat)k).Item) :
        new NumberFormatKind.Decimal();

    private static Fs.Hyperlinks.HyperlinkTarget ToFsHyperlinkTarget(HyperlinkTarget t) => t switch
    {
        HyperlinkTarget.InternalBookmark b => Fs.Hyperlinks.HyperlinkTarget.NewInternalBookmark(b.Name),
        HyperlinkTarget.ExternalUrl u => Fs.Hyperlinks.HyperlinkTarget.NewExternalUrl(u.Url),
        _ => Fs.Hyperlinks.HyperlinkTarget.NewExternalUrl("")
    };

    private static HyperlinkTarget FromFsHyperlinkTarget(Fs.Hyperlinks.HyperlinkTarget t) =>
        t.IsInternalBookmark
            ? new HyperlinkTarget.InternalBookmark(((Fs.Hyperlinks.HyperlinkTarget.InternalBookmark)t).Item)
            : new HyperlinkTarget.ExternalUrl(((Fs.Hyperlinks.HyperlinkTarget.ExternalUrl)t).Item);

    private static Fs.PageSetup.PageSize ToFsPageSize(PageSize p) => p switch
    {
        PageSize.Legal => Fs.PageSetup.PageSize.Legal,
        PageSize.A4 => Fs.PageSetup.PageSize.A4,
        PageSize.A3 => Fs.PageSetup.PageSize.A3,
        PageSize.Other other => Fs.PageSetup.PageSize.NewOtherPageSize(other.Code),
        PageSize.Custom custom => Fs.PageSetup.PageSize.NewCustomPageSize(custom.WidthPoints, custom.HeightPoints),
        _ => Fs.PageSetup.PageSize.Letter
    };

    private static PageSize FromFsPageSize(Fs.PageSetup.PageSize p) =>
        p.IsLegal ? new PageSize.Legal() :
        p.IsA4 ? new PageSize.A4() :
        p.IsA3 ? new PageSize.A3() :
        p.IsOtherPageSize ? new PageSize.Other(((Fs.PageSetup.PageSize.OtherPageSize)p).code) :
        p.IsCustomPageSize ? new PageSize.Custom(((Fs.PageSetup.PageSize.CustomPageSize)p).widthPoints, ((Fs.PageSetup.PageSize.CustomPageSize)p).heightPoints) :
        new PageSize.Letter();

    private static Fs.ContentControls.ContentControlType ToFsContentControlType(ContentControlType t) => t switch
    {
        ContentControlType.PlainText plain => Fs.ContentControls.ContentControlType.NewPlainTextControl(plain.MultiLine),
        ContentControlType.DropDown dropDown => Fs.ContentControls.ContentControlType.NewDropDownControl(ToFsList(dropDown.Items, ToFsPair), dropDown.Editable),
        ContentControlType.Date date => Fs.ContentControls.ContentControlType.NewDateControl(ToOptStruct(date.FullDate), ToOpt(date.Format)),
        ContentControlType.CheckBox checkBox => Fs.ContentControls.ContentControlType.NewCheckBoxControl(
            checkBox.Checked,
            checkBox.CheckedSymbol is { } cs ? ToOpt(ToFsPair(cs)) : FSharpOption<Tuple<string, string>>.None,
            checkBox.UncheckedSymbol is { } us ? ToOpt(ToFsPair(us)) : FSharpOption<Tuple<string, string>>.None),
        _ => Fs.ContentControls.ContentControlType.RichTextControl
    };

    private static ContentControlType FromFsContentControlType(Fs.ContentControls.ContentControlType t)
    {
        if (t.IsPlainTextControl)
            return new ContentControlType.PlainText(((Fs.ContentControls.ContentControlType.PlainTextControl)t).multiLine);

        if (t.IsDropDownControl)
        {
            var dd = (Fs.ContentControls.ContentControlType.DropDownControl)t;
            return new ContentControlType.DropDown(FromFsList(dd.items, FromFsPair), dd.editable);
        }

        if (t.IsDateControl)
        {
            var d = (Fs.ContentControls.ContentControlType.DateControl)t;
            return new ContentControlType.Date(FromOptStruct(d.fullDate), FromOpt(d.format));
        }

        if (t.IsCheckBoxControl)
        {
            var cb = (Fs.ContentControls.ContentControlType.CheckBoxControl)t;
            var checkedSymbol = FromOpt(cb.checkedSymbol);
            var uncheckedSymbol = FromOpt(cb.uncheckedSymbol);
            return new ContentControlType.CheckBox(cb.checked_, checkedSymbol is null ? null : FromFsPair(checkedSymbol), uncheckedSymbol is null ? null : FromFsPair(uncheckedSymbol));
        }

        return new ContentControlType.RichText();
    }

    // ----- Records ----------------------------------------------------------------------

    private static Fs.Styles.Indentation ToFsIndentation(Indentation i) =>
        new(ToOptStruct(i.Left), ToOptStruct(i.Right), ToOptStruct(i.FirstLine), ToOptStruct(i.Hanging));

    private static Indentation FromFsIndentation(Fs.Styles.Indentation i) =>
        new() { Left = FromOptStruct(i.Left), Right = FromOptStruct(i.Right), FirstLine = FromOptStruct(i.FirstLine), Hanging = FromOptStruct(i.Hanging) };

    private static Fs.Styles.BorderSide ToFsBorderSide(BorderSide s) =>
        new(ToFsBorderLineStyle(s.Style), ToOptStruct(s.Width), s.Color is { } c ? ToOpt(ToFsColor(c)) : FSharpOption<Fs.Styles.Color>.None);

    private static BorderSide FromFsBorderSide(Fs.Styles.BorderSide s) =>
        new(FromFsBorderLineStyle(s.Style), FromOptStruct(s.Width), FromOpt(s.Color) is { } c ? FromFsColor(c) : null);

    private static Fs.Styles.BorderStyle ToFsBorderStyle(BorderStyle b) =>
        new(
            b.Left is { } l ? ToOpt(ToFsBorderSide(l)) : FSharpOption<Fs.Styles.BorderSide>.None,
            b.Right is { } r ? ToOpt(ToFsBorderSide(r)) : FSharpOption<Fs.Styles.BorderSide>.None,
            b.Top is { } t ? ToOpt(ToFsBorderSide(t)) : FSharpOption<Fs.Styles.BorderSide>.None,
            b.Bottom is { } bo ? ToOpt(ToFsBorderSide(bo)) : FSharpOption<Fs.Styles.BorderSide>.None);

    private static BorderStyle FromFsBorderStyle(Fs.Styles.BorderStyle b) =>
        new()
        {
            Left = FromOpt(b.Left) is { } l ? FromFsBorderSide(l) : null,
            Right = FromOpt(b.Right) is { } r ? FromFsBorderSide(r) : null,
            Top = FromOpt(b.Top) is { } t ? FromFsBorderSide(t) : null,
            Bottom = FromOpt(b.Bottom) is { } bo ? FromFsBorderSide(bo) : null
        };

    private static Fs.Styles.TabStop ToFsTabStop(TabStop t) => new(t.Position, ToFsTabStopAlignment(t.Alignment), ToFsTabLeader(t.Leader));

    private static TabStop FromFsTabStop(Fs.Styles.TabStop t) => new(t.Position, FromFsTabStopAlignment(t.Alignment), FromFsTabLeader(t.Leader));

    private static Fs.Styles.RunStyle ToFsRunStyle(RunStyle s) =>
        new(
            ToOpt(s.FontFamily), ToOptStruct(s.Size), s.Bold, s.Italic,
            s.Underline is { } u ? ToOpt(ToFsUnderlineStyle(u)) : FSharpOption<Fs.Styles.UnderlineStyle>.None,
            s.Strikethrough,
            s.Color is { } c ? ToOpt(ToFsColor(c)) : FSharpOption<Fs.Styles.Color>.None,
            s.Highlight is { } h ? ToOpt(ToFsHighlightColor(h)) : FSharpOption<Fs.Styles.HighlightColor>.None,
            s.VerticalPosition is { } v ? ToOpt(ToFsVerticalPosition(v)) : FSharpOption<Fs.Styles.VerticalPosition>.None,
            s.SmallCaps, s.AllCaps, s.Hidden);

    private static RunStyle FromFsRunStyle(Fs.Styles.RunStyle s) =>
        new()
        {
            FontFamily = FromOpt(s.FontFamily),
            Size = FromOptStruct(s.Size),
            Bold = s.Bold,
            Italic = s.Italic,
            Underline = FromOpt(s.Underline) is { } u ? FromFsUnderlineStyle(u) : null,
            Strikethrough = s.Strikethrough,
            Color = FromOpt(s.Color) is { } c ? FromFsColor(c) : null,
            Highlight = FromOpt(s.Highlight) is { } h ? FromFsHighlightColor(h) : null,
            VerticalPosition = FromOpt(s.VerticalPosition) is { } v ? FromFsVerticalPosition(v) : null,
            SmallCaps = s.SmallCaps,
            AllCaps = s.AllCaps,
            Hidden = s.Hidden
        };

    private static Fs.Styles.ParagraphFormat ToFsParagraphFormat(ParagraphFormat f) =>
        new(
            f.Alignment is { } a ? ToOpt(ToFsParagraphAlignment(a)) : FSharpOption<Fs.Styles.ParagraphAlignment>.None,
            ToOptStruct(f.SpacingBefore), ToOptStruct(f.SpacingAfter),
            f.LineSpacing is { } ls ? ToOpt(ToFsLineSpacingRule(ls)) : FSharpOption<Fs.Styles.LineSpacingRule>.None,
            f.Indentation is { } ind ? ToOpt(ToFsIndentation(ind)) : FSharpOption<Fs.Styles.Indentation>.None,
            f.KeepWithNext, f.PageBreakBefore,
            f.Borders is { } b ? ToOpt(ToFsBorderStyle(b)) : FSharpOption<Fs.Styles.BorderStyle>.None,
            f.Shading is { } sh ? ToOpt(ToFsColor(sh)) : FSharpOption<Fs.Styles.Color>.None,
            ToFsList(f.TabStops, ToFsTabStop));

    private static ParagraphFormat FromFsParagraphFormat(Fs.Styles.ParagraphFormat f) =>
        new()
        {
            Alignment = FromOpt(f.Alignment) is { } a ? FromFsParagraphAlignment(a) : null,
            SpacingBefore = FromOptStruct(f.SpacingBefore),
            SpacingAfter = FromOptStruct(f.SpacingAfter),
            LineSpacing = FromOpt(f.LineSpacing) is { } ls ? FromFsLineSpacingRule(ls) : null,
            Indentation = FromOpt(f.Indentation) is { } ind ? FromFsIndentation(ind) : null,
            KeepWithNext = f.KeepWithNext,
            PageBreakBefore = f.PageBreakBefore,
            Borders = FromOpt(f.Borders) is { } b ? FromFsBorderStyle(b) : null,
            Shading = FromOpt(f.Shading) is { } sh ? FromFsColor(sh) : null,
            TabStops = FromFsList(f.TabStops, FromFsTabStop)
        };

    private static Fs.NamedStyles.StyleDefinition ToFsStyleDefinition(StyleDefinition d) =>
        new(
            d.Id, d.Name, ToFsStyleTargetType(d.Type), ToOpt(d.BasedOn),
            d.RunFormat is { } rf ? ToOpt(ToFsRunStyle(rf)) : FSharpOption<Fs.Styles.RunStyle>.None,
            d.ParaFormat is { } pf ? ToOpt(ToFsParagraphFormat(pf)) : FSharpOption<Fs.Styles.ParagraphFormat>.None);

    private static StyleDefinition FromFsStyleDefinition(Fs.NamedStyles.StyleDefinition d) =>
        new()
        {
            Id = d.Id,
            Name = d.Name,
            Type = FromFsStyleTargetType(d.Type),
            BasedOn = FromOpt(d.BasedOn),
            RunFormat = FromOpt(d.RunFormat) is { } rf ? FromFsRunStyle(rf) : null,
            ParaFormat = FromOpt(d.ParaFormat) is { } pf ? FromFsParagraphFormat(pf) : null
        };

    private static Fs.Numbering.ListLevel ToFsListLevel(ListLevel l) =>
        new(ToFsNumberFormatKind(l.Format), l.Text, ToOptStruct(l.IndentLeft), ToOptStruct(l.HangingIndent), ToOptStruct(l.StartAt));

    private static ListLevel FromFsListLevel(Fs.Numbering.ListLevel l) =>
        new() { Format = FromFsNumberFormatKind(l.Format), Text = l.Text, IndentLeft = FromOptStruct(l.IndentLeft), HangingIndent = FromOptStruct(l.HangingIndent), StartAt = FromOptStruct(l.StartAt) };

    private static Fs.Numbering.NumberingDefinition ToFsNumberingDefinition(NumberingDefinition d) =>
        new(d.Id, ToFsList(d.Levels, ToFsListLevel));

    private static NumberingDefinition FromFsNumberingDefinition(Fs.Numbering.NumberingDefinition d) =>
        new(d.Id, FromFsList(d.Levels, FromFsListLevel));

    private static Fs.ContentControls.ContentControlProps ToFsContentControlProps(ContentControlProps p) =>
        new(
            ToOpt(p.Alias), ToOpt(p.Tag),
            p.Lock is { } l ? ToOpt(ToFsContentControlLock(l)) : FSharpOption<Fs.ContentControls.ContentControlLock>.None,
            ToFsContentControlType(p.Type));

    private static ContentControlProps FromFsContentControlProps(Fs.ContentControls.ContentControlProps p) =>
        new() { Alias = FromOpt(p.Alias), Tag = FromOpt(p.Tag), Lock = FromOpt(p.Lock) is { } l ? FromFsContentControlLock(l) : null, Type = FromFsContentControlType(p.Type) };

    private static Fs.PageSetup.PageMargins ToFsPageMargins(PageMargins m) =>
        new(m.Top, m.Bottom, m.Left, m.Right, m.Header, m.Footer, m.Gutter);

    private static PageMargins FromFsPageMargins(Fs.PageSetup.PageMargins m) =>
        new() { Top = m.Top, Bottom = m.Bottom, Left = m.Left, Right = m.Right, Header = m.Header, Footer = m.Footer, Gutter = m.Gutter };

    private static Fs.PageSetup.NoteNumberingSettings ToFsNoteNumberingSettings(NoteNumberingSettings s) =>
        new(ToFsNumberFormatKind(s.Format), ToOptStruct(s.StartAt), ToFsNoteNumberRestart(s.Restart));

    private static NoteNumberingSettings FromFsNoteNumberingSettings(Fs.PageSetup.NoteNumberingSettings s) =>
        new() { Format = FromFsNumberFormatKind(s.Format), StartAt = FromOptStruct(s.StartAt), Restart = FromFsNoteNumberRestart(s.Restart) };

    private static Fs.Tables.TableBorders ToFsTableBorders(TableBorders b) =>
        new(
            ToFsBorderStyle(b.Outer),
            b.InsideHorizontal is { } ih ? ToOpt(ToFsBorderSide(ih)) : FSharpOption<Fs.Styles.BorderSide>.None,
            b.InsideVertical is { } iv ? ToOpt(ToFsBorderSide(iv)) : FSharpOption<Fs.Styles.BorderSide>.None);

    private static TableBorders FromFsTableBorders(Fs.Tables.TableBorders b) =>
        new()
        {
            Outer = FromFsBorderStyle(b.Outer),
            InsideHorizontal = FromOpt(b.InsideHorizontal) is { } ih ? FromFsBorderSide(ih) : null,
            InsideVertical = FromOpt(b.InsideVertical) is { } iv ? FromFsBorderSide(iv) : null
        };

    private static Fs.Tables.CellMargins ToFsCellMargins(CellMargins m) =>
        new(ToOptStruct(m.Top), ToOptStruct(m.Bottom), ToOptStruct(m.Left), ToOptStruct(m.Right));

    private static CellMargins FromFsCellMargins(Fs.Tables.CellMargins m) =>
        new() { Top = FromOptStruct(m.Top), Bottom = FromOptStruct(m.Bottom), Left = FromOptStruct(m.Left), Right = FromOptStruct(m.Right) };

    private static Fs.Tables.TableCellProps ToFsTableCellProps(TableCellProps p) =>
        new(
            ToOptStruct(p.GridSpan),
            p.VerticalMerge is { } vm ? ToOpt(ToFsVerticalMergeKind(vm)) : FSharpOption<Fs.Tables.VerticalMergeKind>.None,
            p.Shading is { } sh ? ToOpt(ToFsColor(sh)) : FSharpOption<Fs.Styles.Color>.None,
            p.Borders is { } b ? ToOpt(ToFsTableBorders(b)) : FSharpOption<Fs.Tables.TableBorders>.None,
            ToOptStruct(p.Width),
            p.Margins is { } m ? ToOpt(ToFsCellMargins(m)) : FSharpOption<Fs.Tables.CellMargins>.None);

    private static TableCellProps FromFsTableCellProps(Fs.Tables.TableCellProps p) =>
        new()
        {
            GridSpan = FromOptStruct(p.GridSpan),
            VerticalMerge = FromOpt(p.VerticalMerge) is { } vm ? FromFsVerticalMergeKind(vm) : null,
            Shading = FromOpt(p.Shading) is { } sh ? FromFsColor(sh) : null,
            Borders = FromOpt(p.Borders) is { } b ? FromFsTableBorders(b) : null,
            Width = FromOptStruct(p.Width),
            Margins = FromOpt(p.Margins) is { } m ? FromFsCellMargins(m) : null
        };

    private static Fs.Tables.TableStyleRef ToFsTableStyleRef(TableStyleRef r) =>
        new(r.Name, r.FirstRowBanding, r.LastRowBanding, r.BandedRows, r.BandedColumns);

    private static TableStyleRef FromFsTableStyleRef(Fs.Tables.TableStyleRef r) =>
        new() { Name = r.Name, FirstRowBanding = r.FirstRowBanding, LastRowBanding = r.LastRowBanding, BandedRows = r.BandedRows, BandedColumns = r.BandedColumns };

    private static Fs.Tables.TableStyleRegion ToFsTableStyleRegion(TableStyleRegion r) =>
        new(
            r.RunFormat is { } rf ? ToOpt(ToFsRunStyle(rf)) : FSharpOption<Fs.Styles.RunStyle>.None,
            r.ParaFormat is { } pf ? ToOpt(ToFsParagraphFormat(pf)) : FSharpOption<Fs.Styles.ParagraphFormat>.None,
            r.CellShading is { } cs ? ToOpt(ToFsColor(cs)) : FSharpOption<Fs.Styles.Color>.None);

    private static TableStyleRegion FromFsTableStyleRegion(Fs.Tables.TableStyleRegion r) =>
        new()
        {
            RunFormat = FromOpt(r.RunFormat) is { } rf ? FromFsRunStyle(rf) : null,
            ParaFormat = FromOpt(r.ParaFormat) is { } pf ? FromFsParagraphFormat(pf) : null,
            CellShading = FromOpt(r.CellShading) is { } cs ? FromFsColor(cs) : null
        };

    private static Fs.Tables.TableStyleDefinition ToFsTableStyleDefinition(TableStyleDefinition d) =>
        new(
            d.Id, d.Name, ToOpt(d.BasedOn),
            d.Borders is { } b ? ToOpt(ToFsTableBorders(b)) : FSharpOption<Fs.Tables.TableBorders>.None,
            ToFsTableStyleRegion(d.WholeTable), ToFsTableStyleRegion(d.FirstRow), ToFsTableStyleRegion(d.LastRow),
            ToFsTableStyleRegion(d.FirstColumn), ToFsTableStyleRegion(d.LastColumn),
            ToFsTableStyleRegion(d.BandedRow), ToFsTableStyleRegion(d.BandedColumn),
            ToFsTableStyleRegion(d.BandedRow2), ToFsTableStyleRegion(d.BandedColumn2),
            ToFsTableStyleRegion(d.NorthEastCell), ToFsTableStyleRegion(d.NorthWestCell),
            ToFsTableStyleRegion(d.SouthEastCell), ToFsTableStyleRegion(d.SouthWestCell));

    private static TableStyleDefinition FromFsTableStyleDefinition(Fs.Tables.TableStyleDefinition d) =>
        new()
        {
            Id = d.Id,
            Name = d.Name,
            BasedOn = FromOpt(d.BasedOn),
            Borders = FromOpt(d.Borders) is { } b ? FromFsTableBorders(b) : null,
            WholeTable = FromFsTableStyleRegion(d.WholeTable),
            FirstRow = FromFsTableStyleRegion(d.FirstRow),
            LastRow = FromFsTableStyleRegion(d.LastRow),
            FirstColumn = FromFsTableStyleRegion(d.FirstColumn),
            LastColumn = FromFsTableStyleRegion(d.LastColumn),
            BandedRow = FromFsTableStyleRegion(d.BandedRow),
            BandedColumn = FromFsTableStyleRegion(d.BandedColumn),
            BandedRow2 = FromFsTableStyleRegion(d.BandedRow2),
            BandedColumn2 = FromFsTableStyleRegion(d.BandedColumn2),
            NorthEastCell = FromFsTableStyleRegion(d.NorthEastCell),
            NorthWestCell = FromFsTableStyleRegion(d.NorthWestCell),
            SouthEastCell = FromFsTableStyleRegion(d.SouthEastCell),
            SouthWestCell = FromFsTableStyleRegion(d.SouthWestCell)
        };

    private static Fs.Images.ImageEntry ToFsImageEntry(ImageEntry e) =>
        new(e.Data, ToFsImageFormat(e.Format), e.WidthEmu, e.HeightEmu, ToOpt(e.AltText));

    private static ImageEntry FromFsImageEntry(Fs.Images.ImageEntry e) =>
        new() { Data = e.Data, Format = FromFsImageFormat(e.Format), WidthEmu = e.WidthEmu, HeightEmu = e.HeightEmu, AltText = FromOpt(e.AltText) };

    private static Fs.Protection.DocumentProtection ToFsDocumentProtection(DocumentProtection p) =>
        new(p.Edit is { } e ? ToOpt(ToFsEditRestriction(e)) : FSharpOption<Fs.Protection.EditRestriction>.None, ToOpt(p.Password));

    private static DocumentProtection FromFsDocumentProtection(Fs.Protection.DocumentProtection p) =>
        new() { Edit = FromOpt(p.Edit) is { } e ? FromFsEditRestriction(e) : null, Password = FromOpt(p.Password) };

    private static Fs.DocumentProperties.DocumentProperties ToFsDocumentProperties(DocumentProperties p) =>
        new(ToOpt(p.Title), ToOpt(p.Author), ToOpt(p.Subject), ToOpt(p.Keywords), ToOpt(p.Comments), ToOpt(p.Category), ToOpt(p.Company));

    private static DocumentProperties FromFsDocumentProperties(Fs.DocumentProperties.DocumentProperties p) =>
        new() { Title = FromOpt(p.Title), Author = FromOpt(p.Author), Subject = FromOpt(p.Subject), Keywords = FromOpt(p.Keywords), Comments = FromOpt(p.Comments), Category = FromOpt(p.Category), Company = FromOpt(p.Company) };

    private static Fs.Revisions.Revision ToFsRevision(Revision r) => new(ToFsRevisionKind(r.Kind), r.Author, ToOptStruct(r.Date));

    private static Revision FromFsRevision(Fs.Revisions.Revision r) => new(FromFsRevisionKind(r.Kind), r.Author, FromOptStruct(r.Date));

    // ----- Recursive content: Inline / Block / Paragraph / tables / sections / document -----
    //
    // `Inline`/`Block` are mutually recursive (a footnote's own body is a `Block list`,
    // a table cell's content is a `Block list` too), matching the F# core's own `Model.fs`
    // recursion - these functions call each other the same way.

    private static Fs.Model.Inline ToFsInline(Inline i) => i switch
    {
        Inline.Run run => Fs.Model.Inline.NewRun(run.Text, run.Style is { } s ? ToOpt(ToFsRunStyle(s)) : FSharpOption<Fs.Styles.RunStyle>.None, ToOpt(run.StyleId)),
        Inline.LineBreak => Fs.Model.Inline.LineBreak,
        Inline.Tab => Fs.Model.Inline.Tab,
        Inline.PageBreak => Fs.Model.Inline.PageBreak,
        Inline.Image img => Fs.Model.Inline.NewImage(ToFsImageEntry(img.Entry)),
        Inline.Hyperlink hl => Fs.Model.Inline.NewHyperlink(ToFsHyperlinkTarget(hl.Target), ToFsList(hl.Runs, ToFsInline), ToOpt(hl.Tooltip)),
        Inline.Bookmark bm => Fs.Model.Inline.NewBookmark(bm.Name, ToFsList(bm.Content, ToFsInline)),
        Inline.BookmarkRangeStart brs => Fs.Model.Inline.NewBookmarkRangeStart(brs.Name),
        Inline.BookmarkRangeEnd bre => Fs.Model.Inline.NewBookmarkRangeEnd(bre.Name),
        Inline.Comment c => Fs.Model.Inline.NewComment(c.Author, ToOpt(c.Initials), ToOptStruct(c.Date), c.Text, ToFsList(c.Content, ToFsInline)),
        Inline.CommentRangeStart crs => Fs.Model.Inline.NewCommentRangeStart(crs.Id, crs.Author, ToOpt(crs.Initials), ToOptStruct(crs.Date), crs.Text),
        Inline.CommentRangeEnd cre => Fs.Model.Inline.NewCommentRangeEnd(cre.Id),
        Inline.Field f => Fs.Model.Inline.NewField(f.Instruction, ToOpt(f.CachedResult)),
        Inline.Footnote fn => Fs.Model.Inline.NewFootnote(ToFsList(fn.Content, ToFsBlock)),
        Inline.Endnote en => Fs.Model.Inline.NewEndnote(ToFsList(en.Content, ToFsBlock)),
        Inline.TrackedChange tc => Fs.Model.Inline.NewTrackedChange(ToFsRevision(tc.Revision), ToFsList(tc.Content, ToFsInline)),
        Inline.ContentControl cc => Fs.Model.Inline.NewInlineContentControl(ToFsContentControlProps(cc.Props), ToFsList(cc.Content, ToFsInline)),
        _ => throw new ArgumentOutOfRangeException(nameof(i), i, "Unknown Inline case")
    };

    private static Inline FromFsInline(Fs.Model.Inline i)
    {
        if (i.IsRun)
        {
            var r = (Fs.Model.Inline.Run)i;
            return new Inline.Run(r.text, FromOpt(r.style) is { } s ? FromFsRunStyle(s) : null, FromOpt(r.styleId));
        }

        if (i.IsLineBreak) return new Inline.LineBreak();
        if (i.IsTab) return new Inline.Tab();
        if (i.IsPageBreak) return new Inline.PageBreak();

        if (i.IsImage)
            return new Inline.Image(FromFsImageEntry(((Fs.Model.Inline.Image)i).Item));

        if (i.IsHyperlink)
        {
            var hl = (Fs.Model.Inline.Hyperlink)i;
            return new Inline.Hyperlink(FromFsHyperlinkTarget(hl.target), FromFsList(hl.runs, FromFsInline), FromOpt(hl.tooltip));
        }

        if (i.IsBookmark)
        {
            var bm = (Fs.Model.Inline.Bookmark)i;
            return new Inline.Bookmark(bm.name, FromFsList(bm.content, FromFsInline));
        }

        if (i.IsBookmarkRangeStart)
            return new Inline.BookmarkRangeStart(((Fs.Model.Inline.BookmarkRangeStart)i).name);

        if (i.IsBookmarkRangeEnd)
            return new Inline.BookmarkRangeEnd(((Fs.Model.Inline.BookmarkRangeEnd)i).name);

        if (i.IsComment)
        {
            var c = (Fs.Model.Inline.Comment)i;
            return new Inline.Comment(c.author, FromOpt(c.initials), FromOptStruct(c.date), c.text, FromFsList(c.content, FromFsInline));
        }

        if (i.IsCommentRangeStart)
        {
            var crs = (Fs.Model.Inline.CommentRangeStart)i;
            return new Inline.CommentRangeStart(crs.id, crs.author, FromOpt(crs.initials), FromOptStruct(crs.date), crs.text);
        }

        if (i.IsCommentRangeEnd)
            return new Inline.CommentRangeEnd(((Fs.Model.Inline.CommentRangeEnd)i).id);

        if (i.IsField)
        {
            var f = (Fs.Model.Inline.Field)i;
            return new Inline.Field(f.instruction, FromOpt(f.cachedResult));
        }

        if (i.IsFootnote)
            return new Inline.Footnote(FromFsList(((Fs.Model.Inline.Footnote)i).content, FromFsBlock));

        if (i.IsEndnote)
            return new Inline.Endnote(FromFsList(((Fs.Model.Inline.Endnote)i).content, FromFsBlock));

        if (i.IsTrackedChange)
        {
            var tc = (Fs.Model.Inline.TrackedChange)i;
            return new Inline.TrackedChange(FromFsRevision(tc.revision), FromFsList(tc.content, FromFsInline));
        }

        if (i.IsInlineContentControl)
        {
            var cc = (Fs.Model.Inline.InlineContentControl)i;
            return new Inline.ContentControl(FromFsContentControlProps(cc.props), FromFsList(cc.content, FromFsInline));
        }

        throw new ArgumentOutOfRangeException(nameof(i), i, "Unknown Inline case");
    }

    private static Fs.Model.Block ToFsBlock(Block b) => b switch
    {
        Block.ParagraphBlock pb => Fs.Model.Block.NewParagraphBlock(ToFsParagraph(pb.Para)),
        Block.TableBlock tb => Fs.Model.Block.NewTableBlock(ToFsTableEntry(tb.Entry)),
        Block.ContentControlBlock cc => Fs.Model.Block.NewContentControlBlock(ToFsContentControlProps(cc.Props), ToFsList(cc.Content, ToFsBlock)),
        _ => throw new ArgumentOutOfRangeException(nameof(b), b, "Unknown Block case")
    };

    private static Block FromFsBlock(Fs.Model.Block b)
    {
        if (b.IsParagraphBlock)
            return new Block.ParagraphBlock(FromFsParagraph(((Fs.Model.Block.ParagraphBlock)b).Item));

        if (b.IsTableBlock)
            return new Block.TableBlock(FromFsTableEntry(((Fs.Model.Block.TableBlock)b).Item));

        if (b.IsContentControlBlock)
        {
            var cc = (Fs.Model.Block.ContentControlBlock)b;
            return new Block.ContentControlBlock(FromFsContentControlProps(cc.props), FromFsList(cc.content, FromFsBlock));
        }

        throw new ArgumentOutOfRangeException(nameof(b), b, "Unknown Block case");
    }

    private static Fs.Model.Paragraph ToFsParagraph(Paragraph p) =>
        new(
            ToFsList(p.Inlines, ToFsInline), ToOpt(p.StyleId),
            p.Format is { } f ? ToOpt(ToFsParagraphFormat(f)) : FSharpOption<Fs.Styles.ParagraphFormat>.None,
            p.Numbering is { } n ? ToOpt(Tuple.Create(n.NumId, n.Level)) : FSharpOption<Tuple<int, int>>.None,
            p.MarkRevision is { } mr ? ToOpt(ToFsRevision(mr)) : FSharpOption<Fs.Revisions.Revision>.None);

    private static Paragraph FromFsParagraph(Fs.Model.Paragraph p) =>
        new()
        {
            Inlines = FromFsList(p.Inlines, FromFsInline),
            StyleId = FromOpt(p.StyleId),
            Format = FromOpt(p.Format) is { } f ? FromFsParagraphFormat(f) : null,
            Numbering = FromOpt(p.Numbering) is { } n ? (n.Item1, n.Item2) : null,
            MarkRevision = FromOpt(p.MarkRevision) is { } mr ? FromFsRevision(mr) : null
        };

    private static Fs.Model.TableCell ToFsTableCell(TableCell c) => new(ToFsList(c.Content, ToFsBlock), ToFsTableCellProps(c.Props));

    private static TableCell FromFsTableCell(Fs.Model.TableCell c) => new() { Content = FromFsList(c.Content, FromFsBlock), Props = FromFsTableCellProps(c.Props) };

    private static Fs.Model.TableRow ToFsTableRow(TableRow r) => new(ToFsList(r.Cells, ToFsTableCell), ToOptStruct(r.Height), r.RepeatAsHeader);

    private static TableRow FromFsTableRow(Fs.Model.TableRow r) => new() { Cells = FromFsList(r.Cells, FromFsTableCell), Height = FromOptStruct(r.Height), RepeatAsHeader = r.RepeatAsHeader };

    private static Fs.Model.TableEntry ToFsTableEntry(TableEntry e) =>
        new(
            ToFsList(e.Rows, ToFsTableRow), ListModule.OfSeq(e.ColumnWidths),
            e.Style is { } s ? ToOpt(ToFsTableStyleRef(s)) : FSharpOption<Fs.Tables.TableStyleRef>.None,
            e.Borders is { } b ? ToOpt(ToFsTableBorders(b)) : FSharpOption<Fs.Tables.TableBorders>.None,
            e.CellMargins is { } m ? ToOpt(ToFsCellMargins(m)) : FSharpOption<Fs.Tables.CellMargins>.None);

    private static TableEntry FromFsTableEntry(Fs.Model.TableEntry e) =>
        new()
        {
            Rows = FromFsList(e.Rows, FromFsTableRow),
            ColumnWidths = e.ColumnWidths.ToList(),
            Style = FromOpt(e.Style) is { } s ? FromFsTableStyleRef(s) : null,
            Borders = FromOpt(e.Borders) is { } b ? FromFsTableBorders(b) : null,
            CellMargins = FromOpt(e.CellMargins) is { } m ? FromFsCellMargins(m) : null
        };

    private static Fs.Model.HeaderFooterSet ToFsHeaderFooterSet(HeaderFooterSet h) =>
        new(
            h.Default is { } d ? ToOpt(ToFsList(d, ToFsBlock)) : FSharpOption<FSharpList<Fs.Model.Block>>.None,
            h.First is { } f ? ToOpt(ToFsList(f, ToFsBlock)) : FSharpOption<FSharpList<Fs.Model.Block>>.None,
            h.Even is { } e ? ToOpt(ToFsList(e, ToFsBlock)) : FSharpOption<FSharpList<Fs.Model.Block>>.None);

    private static HeaderFooterSet FromFsHeaderFooterSet(Fs.Model.HeaderFooterSet h) =>
        new()
        {
            Default = FromOpt(h.Default) is { } d ? FromFsList(d, FromFsBlock) : null,
            First = FromOpt(h.First) is { } f ? FromFsList(f, FromFsBlock) : null,
            Even = FromOpt(h.Even) is { } e ? FromFsList(e, FromFsBlock) : null
        };

    private static Fs.Model.SectionProperties ToFsSectionProperties(SectionProperties p) =>
        new(
            ToFsPageSize(p.PageSize), ToFsPageOrientation(p.Orientation), ToFsPageMargins(p.Margins),
            p.Header is { } h ? ToOpt(ToFsHeaderFooterSet(h)) : FSharpOption<Fs.Model.HeaderFooterSet>.None,
            p.Footer is { } f ? ToOpt(ToFsHeaderFooterSet(f)) : FSharpOption<Fs.Model.HeaderFooterSet>.None,
            ToOptStruct(p.PageNumberStart), p.Columns, ToFsSectionBreakType(p.BreakType),
            p.FootnoteNumbering is { } fn ? ToOpt(ToFsNoteNumberingSettings(fn)) : FSharpOption<Fs.PageSetup.NoteNumberingSettings>.None,
            p.EndnoteNumbering is { } en ? ToOpt(ToFsNoteNumberingSettings(en)) : FSharpOption<Fs.PageSetup.NoteNumberingSettings>.None);

    private static SectionProperties FromFsSectionProperties(Fs.Model.SectionProperties p) =>
        new()
        {
            PageSize = FromFsPageSize(p.PageSize),
            Orientation = FromFsPageOrientation(p.Orientation),
            Margins = FromFsPageMargins(p.Margins),
            Header = FromOpt(p.Header) is { } h ? FromFsHeaderFooterSet(h) : null,
            Footer = FromOpt(p.Footer) is { } f ? FromFsHeaderFooterSet(f) : null,
            PageNumberStart = FromOptStruct(p.PageNumberStart),
            Columns = p.Columns,
            BreakType = FromFsSectionBreakType(p.BreakType),
            FootnoteNumbering = FromOpt(p.FootnoteNumbering) is { } fn ? FromFsNoteNumberingSettings(fn) : null,
            EndnoteNumbering = FromOpt(p.EndnoteNumbering) is { } en ? FromFsNoteNumberingSettings(en) : null
        };

    private static Fs.Model.Section ToFsSection(Section s) => new(ToFsList(s.Body, ToFsBlock), ToFsSectionProperties(s.Properties));

    private static Section FromFsSection(Fs.Model.Section s) => new(FromFsList(s.Body, FromFsBlock), FromFsSectionProperties(s.Properties));

    // ----- Top level ------------------------------------------------------------------------

    public static Fs.Model.Document ToFSharp(Document d) =>
        new(
            ToFsList(d.Sections, ToFsSection),
            ToFsList(d.Styles, ToFsStyleDefinition),
            ToFsList(d.Numbering, ToFsNumberingDefinition),
            d.Protection is { } p ? ToOpt(ToFsDocumentProtection(p)) : FSharpOption<Fs.Protection.DocumentProtection>.None,
            ToOpt(d.VbaProject),
            ToFsDocumentProperties(d.Properties),
            ToFsList(d.TableStyles, ToFsTableStyleDefinition));

    public static Document FromFSharp(Fs.Model.Document d) =>
        new()
        {
            Sections = FromFsList(d.Sections, FromFsSection),
            Styles = FromFsList(d.Styles, FromFsStyleDefinition),
            Numbering = FromFsList(d.Numbering, FromFsNumberingDefinition),
            Protection = FromOpt(d.Protection) is { } p ? FromFsDocumentProtection(p) : null,
            VbaProject = FromOpt(d.VbaProject),
            Properties = FromFsDocumentProperties(d.Properties),
            TableStyles = FromFsList(d.TableStyles, FromFsTableStyleDefinition)
        };
}
