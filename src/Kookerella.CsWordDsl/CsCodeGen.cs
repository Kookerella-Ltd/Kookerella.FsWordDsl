using System.Globalization;
using System.Text;

namespace Kookerella.CsWordDsl;

/// <summary>
/// Renders a <see cref="Document"/> back out as a self-contained C# file that regenerates
/// an equivalent file when run - the reverse of <see cref="DocumentIO.Load(string)"/> one
/// level further: loading turns a file into this wrapper's types, this turns those types
/// into C# *source text*. Every renderer below is a direct, mechanical mirror of a type's
/// own fluent API (diffing against <c>.Default</c>/<see langword="null"/> where one exists,
/// so generated code only mentions what isn't already implied) - there's no separate
/// "codegen model", just string-building over this assembly's own public types. Mirrors the
/// sibling Kookerella.CsOpenXmlDsl project's own <c>CsCodeGen</c>.
/// <para>
/// The emitted file targets .NET's "file-based apps" feature (<c>dotnet run script.cs</c>)
/// rather than a traditional project - <see cref="Generate"/>'s own <c>referenceLines</c>
/// parameter is whatever raw <c>#:package</c>/<c>#:project</c> directives the caller needs
/// so the file can locate this assembly; this class has no opinion on that.
/// </para>
/// </summary>
public static class CsCodeGen
{
    private static string RenderString(string s) =>
        "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t") + "\"";

    private static string? RenderStringOrNull(string? s) => s is null ? null : RenderString(s);

    private static string RenderDouble(double d)
    {
        if (double.IsNaN(d)) return "double.NaN";
        if (double.IsPositiveInfinity(d)) return "double.PositiveInfinity";
        if (double.IsNegativeInfinity(d)) return "double.NegativeInfinity";
        var s = d.ToString(CultureInfo.InvariantCulture);
        return s.Contains('.') || s.Contains('E') ? s : s + ".0";
    }

    private static string RenderBool(bool b) => b ? "true" : "false";

    private static string RenderDateTime(DateTime d) => $"new DateTime({d.Ticks}L, DateTimeKind.{d.Kind})";

    private static string RenderChar(char c) => c switch
    {
        '\'' => "'\\''",
        '\\' => "'\\\\'",
        _ when c < 32 || c > 126 => $"'\\u{(int)c:x4}'",
        _ => $"'{c}'"
    };

    private static string RenderByteArray(byte[] bytes) => $"Convert.FromBase64String({RenderString(Convert.ToBase64String(bytes))})";

    private static string RenderPair((string A, string B) p) => $"({RenderString(p.A)}, {RenderString(p.B)})";

    // ----- Simple values ---------------------------------------------------------------------

    private static string RenderRgb(Color.Rgb rgb) => $"new Color.Rgb({rgb.R}, {rgb.G}, {rgb.B})";

    private static string RenderColor(Color c) => c switch
    {
        Color.Rgb rgb => RenderRgb(rgb),
        Color.Auto => "new Color.Auto()",
        Color.Theme theme => $"new Color.Theme(ThemeColorKind.{theme.Kind}, {RenderRgb(theme.Fallback)}, {(theme.Tint is { } t ? RenderDouble(t) : "null")}, {(theme.Shade is { } sh ? RenderDouble(sh) : "null")})",
        _ => "new Color.Auto()"
    };

    private static string RenderUnderlineStyle(UnderlineStyle u) => u switch
    {
        UnderlineStyle.Other other => $"new UnderlineStyle.Other({RenderString(other.Raw)})",
        _ => $"new UnderlineStyle.{u.GetType().Name}()"
    };

    private static string RenderBorderLineStyle(BorderLineStyle b) => b switch
    {
        BorderLineStyle.Other other => $"new BorderLineStyle.Other({RenderString(other.Raw)})",
        _ => $"new BorderLineStyle.{b.GetType().Name}()"
    };

    private static string RenderTabStopAlignment(TabStopAlignment a) => a switch
    {
        TabStopAlignment.Other other => $"new TabStopAlignment.Other({RenderString(other.Raw)})",
        _ => $"new TabStopAlignment.{a.GetType().Name}()"
    };

    private static string RenderLineSpacingRule(LineSpacingRule r) => r switch
    {
        LineSpacingRule.AtLeast v => $"new LineSpacingRule.AtLeast({RenderDouble(v.Points)})",
        LineSpacingRule.Exactly v => $"new LineSpacingRule.Exactly({RenderDouble(v.Points)})",
        LineSpacingRule.Multiple v => $"new LineSpacingRule.Multiple({RenderDouble(v.Factor)})",
        _ => $"new LineSpacingRule.{r.GetType().Name}()"
    };

    private static string RenderNumberFormatKind(NumberFormatKind k) => k switch
    {
        NumberFormatKind.Bullet b => $"new NumberFormatKind.Bullet({RenderChar(b.Glyph)}, {RenderString(b.FontFamily)})",
        NumberFormatKind.Other o => $"new NumberFormatKind.Other({RenderString(o.Raw)})",
        _ => $"new NumberFormatKind.{k.GetType().Name}()"
    };

    private static string RenderHyperlinkTarget(HyperlinkTarget t) => t switch
    {
        HyperlinkTarget.ExternalUrl u => $"new HyperlinkTarget.ExternalUrl({RenderString(u.Url)})",
        HyperlinkTarget.InternalBookmark b => $"new HyperlinkTarget.InternalBookmark({RenderString(b.Name)})",
        _ => "new HyperlinkTarget.ExternalUrl(\"\")"
    };

    private static string RenderPageSize(PageSize p) => p switch
    {
        PageSize.Other o => $"new PageSize.Other({o.Code})",
        PageSize.Custom c => $"new PageSize.Custom({RenderDouble(c.WidthPoints)}, {RenderDouble(c.HeightPoints)})",
        _ => $"new PageSize.{p.GetType().Name}()"
    };

    private static string RenderContentControlType(ContentControlType t) => t switch
    {
        ContentControlType.PlainText p => $"new ContentControlType.PlainText({RenderBool(p.MultiLine)})",
        ContentControlType.RichText => "new ContentControlType.RichText()",
        ContentControlType.DropDown d => $"new ContentControlType.DropDown([{string.Join(", ", d.Items.Select(RenderPair))}], {RenderBool(d.Editable)})",
        ContentControlType.Date dt => $"new ContentControlType.Date({(dt.FullDate is { } fd ? RenderDateTime(fd) : "null")}, {RenderStringOrNull(dt.Format) ?? "null"})",
        ContentControlType.CheckBox cb => $"new ContentControlType.CheckBox({RenderBool(cb.Checked)}, {(cb.CheckedSymbol is { } cs ? RenderPair(cs) : "null")}, {(cb.UncheckedSymbol is { } us ? RenderPair(us) : "null")})",
        _ => "new ContentControlType.RichText()"
    };

    // ----- Records ----------------------------------------------------------------------

    private static string RenderRunStyle(RunStyle s)
    {
        var sb = new StringBuilder("RunStyle.Default");
        if (s.FontFamily is { } ff) sb.Append($".WithFontFamily({RenderString(ff)})");
        if (s.Size is { } sz) sb.Append($".WithSize({RenderDouble(sz)})");
        if (s.Bold) sb.Append(".AsBold()");
        if (s.Italic) sb.Append(".AsItalic()");
        if (s.Underline is { } u) sb.Append($".WithUnderline({RenderUnderlineStyle(u)})");
        if (s.Strikethrough) sb.Append(".AsStrikethrough()");
        if (s.Color is { } c) sb.Append($".WithColor({RenderColor(c)})");
        if (s.Highlight is { } h) sb.Append($".WithHighlight(HighlightColor.{h})");
        if (s.VerticalPosition is { } vp) sb.Append($".WithVerticalPosition(VerticalPosition.{vp})");
        if (s.SmallCaps) sb.Append(".AsSmallCaps()");
        if (s.AllCaps) sb.Append(".AsAllCaps()");
        if (s.Hidden) sb.Append(".AsHidden()");
        return sb.ToString();
    }

    private static string RenderIndentation(Indentation i)
    {
        var sb = new StringBuilder("Indentation.None");
        if (i.Left is { } l) sb.Append($".WithLeft({RenderDouble(l)})");
        if (i.Right is { } r) sb.Append($".WithRight({RenderDouble(r)})");
        if (i.FirstLine is { } fl) sb.Append($".WithFirstLine({RenderDouble(fl)})");
        if (i.Hanging is { } h) sb.Append($".WithHanging({RenderDouble(h)})");
        return sb.ToString();
    }

    private static string RenderBorderSide(BorderSide s)
    {
        var args = new List<string> { RenderBorderLineStyle(s.Style) };
        if (s.Width is { } w) args.Add($"Width: {RenderDouble(w)}");
        if (s.Color is { } c) args.Add($"Color: {RenderColor(c)}");
        return $"new BorderSide({string.Join(", ", args)})";
    }

    private static string RenderBorderStyle(BorderStyle b)
    {
        var sb = new StringBuilder("BorderStyle.None");
        if (b.Left is { } l) sb.Append($".WithLeft({RenderBorderSide(l)})");
        if (b.Right is { } r) sb.Append($".WithRight({RenderBorderSide(r)})");
        if (b.Top is { } t) sb.Append($".WithTop({RenderBorderSide(t)})");
        if (b.Bottom is { } bo) sb.Append($".WithBottom({RenderBorderSide(bo)})");
        return sb.ToString();
    }

    private static string RenderTabStop(TabStop t) => $"new TabStop({RenderDouble(t.Position)}, {RenderTabStopAlignment(t.Alignment)}, TabLeader.{t.Leader})";

    private static string RenderParagraphFormat(ParagraphFormat f)
    {
        var sb = new StringBuilder("ParagraphFormat.Default");
        if (f.Alignment is { } a) sb.Append($".WithAlignment(ParagraphAlignment.{a})");
        if (f.SpacingBefore is { } sb2) sb.Append($".WithSpacingBefore({RenderDouble(sb2)})");
        if (f.SpacingAfter is { } sa) sb.Append($".WithSpacingAfter({RenderDouble(sa)})");
        if (f.LineSpacing is { } ls) sb.Append($".WithLineSpacing({RenderLineSpacingRule(ls)})");
        if (f.Indentation is { } ind) sb.Append($".WithIndentation({RenderIndentation(ind)})");
        if (f.KeepWithNext) sb.Append(".AsKeepWithNext()");
        if (f.PageBreakBefore) sb.Append(".AsPageBreakBefore()");
        if (f.Borders is { } bd) sb.Append($".WithBorders({RenderBorderStyle(bd)})");
        if (f.Shading is { } sh) sb.Append($".WithShading({RenderColor(sh)})");
        if (f.TabStops.Count > 0) sb.Append($".WithTabStops({string.Join(", ", f.TabStops.Select(RenderTabStop))})");
        return sb.ToString();
    }

    private static string RenderContentControlProps(ContentControlProps p)
    {
        var sb = new StringBuilder($"new ContentControlProps {{ Type = {RenderContentControlType(p.Type)}");
        if (p.Alias is { } a) sb.Append($", Alias = {RenderString(a)}");
        if (p.Tag is { } t) sb.Append($", Tag = {RenderString(t)}");
        if (p.Lock is { } l) sb.Append($", Lock = ContentControlLock.{l}");
        sb.Append(" }");
        return sb.ToString();
    }

    private static string RenderImageEntry(ImageEntry e) =>
        $"ImageEntry.FromBytes({RenderByteArray(e.Data)}, ImageFormat.{e.Format}, {e.WidthEmu}L, {e.HeightEmu}L{(e.AltText is { } alt ? $", {RenderString(alt)}" : "")})";

    private static string RenderRevision(Revision r) =>
        $"new Revision(RevisionKind.{r.Kind}, {RenderString(r.Author)}{(r.Date is { } d ? $", {RenderDateTime(d)}" : "")})";

    // ----- Recursive content: Inline / Block / Paragraph / tables / sections ----------------

    private static string RenderInline(Inline i) => i switch
    {
        Inline.Run r => $"new Inline.Run({RenderString(r.Text)}{(r.Style is { } s ? $", {RenderRunStyle(s)}" : "")}{(r.StyleId is { } sid ? $", StyleId: {RenderString(sid)}" : "")})",
        Inline.LineBreak => "new Inline.LineBreak()",
        Inline.Tab => "new Inline.Tab()",
        Inline.PageBreak => "new Inline.PageBreak()",
        Inline.Image img => $"new Inline.Image({RenderImageEntry(img.Entry)})",
        Inline.Hyperlink hl => $"new Inline.Hyperlink({RenderHyperlinkTarget(hl.Target)}, [{string.Join(", ", hl.Runs.Select(RenderInline))}]{(hl.Tooltip is { } tt ? $", {RenderString(tt)}" : "")})",
        Inline.Bookmark bm => $"new Inline.Bookmark({RenderString(bm.Name)}, [{string.Join(", ", bm.Content.Select(RenderInline))}])",
        Inline.BookmarkRangeStart brs => $"new Inline.BookmarkRangeStart({RenderString(brs.Name)})",
        Inline.BookmarkRangeEnd bre => $"new Inline.BookmarkRangeEnd({RenderString(bre.Name)})",
        Inline.Comment c => $"new Inline.Comment({RenderString(c.Author)}, {RenderStringOrNull(c.Initials) ?? "null"}, {(c.Date is { } d ? RenderDateTime(d) : "null")}, {RenderString(c.Text)}, [{string.Join(", ", c.Content.Select(RenderInline))}])",
        Inline.CommentRangeStart crs => $"new Inline.CommentRangeStart({RenderString(crs.Id)}, {RenderString(crs.Author)}, {RenderStringOrNull(crs.Initials) ?? "null"}, {(crs.Date is { } d ? RenderDateTime(d) : "null")}, {RenderString(crs.Text)})",
        Inline.CommentRangeEnd cre => $"new Inline.CommentRangeEnd({RenderString(cre.Id)})",
        Inline.Field f => $"new Inline.Field({RenderString(f.Instruction)}{(f.CachedResult is { } cr ? $", {RenderString(cr)}" : "")})",
        Inline.Footnote fn => $"new Inline.Footnote([{string.Join(", ", fn.Content.Select(RenderBlock))}])",
        Inline.Endnote en => $"new Inline.Endnote([{string.Join(", ", en.Content.Select(RenderBlock))}])",
        Inline.TrackedChange tc => $"new Inline.TrackedChange({RenderRevision(tc.Revision)}, [{string.Join(", ", tc.Content.Select(RenderInline))}])",
        Inline.ContentControl cc => $"new Inline.ContentControl({RenderContentControlProps(cc.Props)}, [{string.Join(", ", cc.Content.Select(RenderInline))}])",
        _ => throw new ArgumentOutOfRangeException(nameof(i), i, "Unknown Inline case")
    };

    private static string RenderParagraph(Paragraph p)
    {
        var args = new List<string> { $"[{string.Join(", ", p.Inlines.Select(RenderInline))}]" };
        if (p.StyleId is { } sid) args.Add($"styleId: {RenderString(sid)}");
        if (p.Format is { } f) args.Add($"format: {RenderParagraphFormat(f)}");
        if (p.Numbering is { } n) args.Add($"numbering: ({n.NumId}, {n.Level})");
        if (p.MarkRevision is { } mr) args.Add($"markRevision: {RenderRevision(mr)}");
        return $"Block.Paragraph({string.Join(", ", args)})";
    }

    private static string RenderTableCellProps(TableCellProps p)
    {
        var sb = new StringBuilder("TableCellProps.Default");
        if (p.GridSpan is { } gs) sb.Append($".WithGridSpan({gs})");
        if (p.VerticalMerge is { } vm) sb.Append($".WithVerticalMerge(VerticalMergeKind.{vm})");
        if (p.Shading is { } sh) sb.Append($".WithShading({RenderColor(sh)})");
        if (p.Borders is { } b) sb.Append($".WithBorders({RenderTableBorders(b)})");
        if (p.Width is { } w) sb.Append($".WithWidth({RenderDouble(w)})");
        if (p.Margins is { } m) sb.Append($".WithMargins({RenderCellMargins(m)})");
        return sb.ToString();
    }

    private static string RenderCellMargins(CellMargins m)
    {
        var parts = new List<string>();
        if (m.Top is { } t) parts.Add($"Top = {RenderDouble(t)}");
        if (m.Bottom is { } b) parts.Add($"Bottom = {RenderDouble(b)}");
        if (m.Left is { } l) parts.Add($"Left = {RenderDouble(l)}");
        if (m.Right is { } r) parts.Add($"Right = {RenderDouble(r)}");
        return parts.Count == 0 ? "CellMargins.Default" : $"new CellMargins {{ {string.Join(", ", parts)} }}";
    }

    private static string RenderTableBorders(TableBorders b)
    {
        var sb = new StringBuilder("TableBorders.None");
        if (b.Outer != BorderStyle.None) sb.Append($".WithOuter({RenderBorderStyle(b.Outer)})");
        if (b.InsideHorizontal is { } ih) sb.Append($".WithInsideHorizontal({RenderBorderSide(ih)})");
        if (b.InsideVertical is { } iv) sb.Append($".WithInsideVertical({RenderBorderSide(iv)})");
        return sb.ToString();
    }

    private static string RenderTableStyleRef(TableStyleRef r)
    {
        var sb = new StringBuilder($"TableStyleRef.Named({RenderString(r.Name)})");
        var flags = new List<string>();
        if (r.FirstRowBanding) flags.Add("FirstRowBanding = true");
        if (r.LastRowBanding) flags.Add("LastRowBanding = true");
        if (r.BandedRows) flags.Add("BandedRows = true");
        if (r.BandedColumns) flags.Add("BandedColumns = true");
        return flags.Count == 0 ? sb.ToString() : $"({sb} with {{ {string.Join(", ", flags)} }})";
    }

    private static string RenderTableCell(TableCell c)
    {
        var content = $"[{string.Join(", ", c.Content.Select(RenderBlock))}]";
        return c.Props == TableCellProps.Default ? $"TableCell.Of({content})" : $"TableCell.Of({content}, {RenderTableCellProps(c.Props)})";
    }

    private static string RenderTableRow(TableRow r)
    {
        var args = new List<string> { $"[{string.Join(", ", r.Cells.Select(RenderTableCell))}]" };
        if (r.Height is { } h) args.Add($"height: {RenderDouble(h)}");
        if (r.RepeatAsHeader) args.Add("repeatAsHeader: true");
        return $"TableRow.Of({string.Join(", ", args)})";
    }

    private static string RenderTable(TableEntry e)
    {
        var args = new List<string>
        {
            $"[{string.Join(", ", e.Rows.Select(RenderTableRow))}]",
            $"[{string.Join(", ", e.ColumnWidths.Select(RenderDouble))}]"
        };
        if (e.Style is { } s) args.Add($"style: {RenderTableStyleRef(s)}");
        if (e.Borders is { } b) args.Add($"borders: {RenderTableBorders(b)}");
        if (e.CellMargins is { } m) args.Add($"cellMargins: {RenderCellMargins(m)}");
        return $"Block.Table({string.Join(", ", args)})";
    }

    private static string RenderBlock(Block b) => b switch
    {
        Block.ParagraphBlock p => RenderParagraph(p.Para),
        Block.TableBlock t => RenderTable(t.Entry),
        Block.ContentControlBlock cc => $"new Block.ContentControlBlock({RenderContentControlProps(cc.Props)}, [{string.Join(", ", cc.Content.Select(RenderBlock))}])",
        _ => throw new ArgumentOutOfRangeException(nameof(b), b, "Unknown Block case")
    };

    private static string RenderHeaderFooterSet(HeaderFooterSet h)
    {
        var sb = new StringBuilder("HeaderFooterSet.None");
        if (h.Default is { } d) sb.Append($".WithDefault([{string.Join(", ", d.Select(RenderBlock))}])");
        if (h.First is { } f) sb.Append($".WithFirst([{string.Join(", ", f.Select(RenderBlock))}])");
        if (h.Even is { } e) sb.Append($".WithEven([{string.Join(", ", e.Select(RenderBlock))}])");
        return sb.ToString();
    }

    private static string RenderNoteNumberingSettings(NoteNumberingSettings s) =>
        $"new NoteNumberingSettings {{ Format = {RenderNumberFormatKind(s.Format)}{(s.StartAt is { } sa ? $", StartAt = {sa}" : "")}, Restart = NoteNumberRestart.{s.Restart} }}";

    private static string RenderSectionProperties(SectionProperties p)
    {
        var sb = new StringBuilder("SectionProperties.Default");
        if (p.PageSize is not PageSize.Letter) sb.Append($".WithPageSize({RenderPageSize(p.PageSize)})");
        if (p.Orientation != PageOrientation.Portrait) sb.Append($".WithOrientation(PageOrientation.{p.Orientation})");
        if (p.Margins != PageMargins.Default) sb.Append($".WithMargins({RenderPageMargins(p.Margins)})");
        if (p.Header is { } h) sb.Append($".WithHeader({RenderHeaderFooterSet(h)})");
        if (p.Footer is { } f) sb.Append($".WithFooter({RenderHeaderFooterSet(f)})");
        if (p.PageNumberStart is { } pns) sb.Append($".WithPageNumberStart({pns})");
        if (p.Columns != 1) sb.Append($".WithColumns({p.Columns})");
        if (p.BreakType != SectionBreakType.NextPage) sb.Append($".WithBreakType(SectionBreakType.{p.BreakType})");
        if (p.FootnoteNumbering is { } fn) sb.Append($".WithFootnoteNumbering({RenderNoteNumberingSettings(fn)})");
        if (p.EndnoteNumbering is { } en) sb.Append($".WithEndnoteNumbering({RenderNoteNumberingSettings(en)})");
        return sb.ToString();
    }

    private static string RenderPageMargins(PageMargins m) =>
        $"new PageMargins {{ Top = {RenderDouble(m.Top)}, Bottom = {RenderDouble(m.Bottom)}, Left = {RenderDouble(m.Left)}, Right = {RenderDouble(m.Right)}, Header = {RenderDouble(m.Header)}, Footer = {RenderDouble(m.Footer)}, Gutter = {RenderDouble(m.Gutter)} }}";

    private static string RenderSection(Section s) =>
        s.Properties == SectionProperties.Default
            ? $"Section.Of([{string.Join(", ", s.Body.Select(RenderBlock))}])"
            : $"Section.With({RenderSectionProperties(s.Properties)}, [{string.Join(", ", s.Body.Select(RenderBlock))}])";

    // ----- Styles / numbering / protection / properties --------------------------------------

    private static string RenderStyleDefinition(StyleDefinition d)
    {
        var sb = new StringBuilder($"new StyleDefinition {{ Id = {RenderString(d.Id)}, Name = {RenderString(d.Name)}, Type = StyleTargetType.{d.Type}");
        if (d.BasedOn is { } b) sb.Append($", BasedOn = {RenderString(b)}");
        if (d.RunFormat is { } rf) sb.Append($", RunFormat = {RenderRunStyle(rf)}");
        if (d.ParaFormat is { } pf) sb.Append($", ParaFormat = {RenderParagraphFormat(pf)}");
        sb.Append(" }");
        return sb.ToString();
    }

    private static string RenderListLevel(ListLevel l)
    {
        var sb = new StringBuilder($"new ListLevel {{ Format = {RenderNumberFormatKind(l.Format)}, Text = {RenderString(l.Text)}");
        if (l.IndentLeft is { } il) sb.Append($", IndentLeft = {RenderDouble(il)}");
        if (l.HangingIndent is { } hi) sb.Append($", HangingIndent = {RenderDouble(hi)}");
        if (l.StartAt is { } sa) sb.Append($", StartAt = {sa}");
        sb.Append(" }");
        return sb.ToString();
    }

    private static string RenderNumberingDefinition(NumberingDefinition d) =>
        $"new NumberingDefinition({d.Id}, [{string.Join(", ", d.Levels.Select(RenderListLevel))}])";

    private static string RenderDocumentProtection(DocumentProtection p)
    {
        var sb = new StringBuilder("new DocumentProtection {");
        var parts = new List<string>();
        if (p.Edit is { } e) parts.Add($"Edit = EditRestriction.{e}");
        if (p.Password is { } pw) parts.Add($"Password = {RenderString(pw)}");
        sb.Append(' ').Append(string.Join(", ", parts)).Append(" }");
        return sb.ToString();
    }

    private static string RenderTableStyleRegion(TableStyleRegion r)
    {
        var sb = new StringBuilder("TableStyleRegion.None");
        if (r.RunFormat is { } rf) sb.Append($".WithRunFormat({RenderRunStyle(rf)})");
        if (r.ParaFormat is { } pf) sb.Append($".WithParaFormat({RenderParagraphFormat(pf)})");
        if (r.CellShading is { } cs) sb.Append($".WithCellShading({RenderColor(cs)})");
        return sb.ToString();
    }

    private static string RenderTableStyleDefinition(TableStyleDefinition d)
    {
        var sb = new StringBuilder($"new TableStyleDefinition {{ Id = {RenderString(d.Id)}, Name = {RenderString(d.Name)}");
        if (d.BasedOn is { } b) sb.Append($", BasedOn = {RenderString(b)}");
        if (d.Borders is { } bo) sb.Append($", Borders = {RenderTableBorders(bo)}");
        if (d.WholeTable != TableStyleRegion.None) sb.Append($", WholeTable = {RenderTableStyleRegion(d.WholeTable)}");
        if (d.FirstRow != TableStyleRegion.None) sb.Append($", FirstRow = {RenderTableStyleRegion(d.FirstRow)}");
        if (d.LastRow != TableStyleRegion.None) sb.Append($", LastRow = {RenderTableStyleRegion(d.LastRow)}");
        if (d.FirstColumn != TableStyleRegion.None) sb.Append($", FirstColumn = {RenderTableStyleRegion(d.FirstColumn)}");
        if (d.LastColumn != TableStyleRegion.None) sb.Append($", LastColumn = {RenderTableStyleRegion(d.LastColumn)}");
        if (d.BandedRow != TableStyleRegion.None) sb.Append($", BandedRow = {RenderTableStyleRegion(d.BandedRow)}");
        if (d.BandedColumn != TableStyleRegion.None) sb.Append($", BandedColumn = {RenderTableStyleRegion(d.BandedColumn)}");
        if (d.BandedRow2 != TableStyleRegion.None) sb.Append($", BandedRow2 = {RenderTableStyleRegion(d.BandedRow2)}");
        if (d.BandedColumn2 != TableStyleRegion.None) sb.Append($", BandedColumn2 = {RenderTableStyleRegion(d.BandedColumn2)}");
        if (d.NorthEastCell != TableStyleRegion.None) sb.Append($", NorthEastCell = {RenderTableStyleRegion(d.NorthEastCell)}");
        if (d.NorthWestCell != TableStyleRegion.None) sb.Append($", NorthWestCell = {RenderTableStyleRegion(d.NorthWestCell)}");
        if (d.SouthEastCell != TableStyleRegion.None) sb.Append($", SouthEastCell = {RenderTableStyleRegion(d.SouthEastCell)}");
        if (d.SouthWestCell != TableStyleRegion.None) sb.Append($", SouthWestCell = {RenderTableStyleRegion(d.SouthWestCell)}");
        sb.Append(" }");
        return sb.ToString();
    }

    private static string RenderDocumentProperties(DocumentProperties p)
    {
        if (p == DocumentProperties.Default)
            return "DocumentProperties.Default";

        var sb = new StringBuilder("DocumentProperties.Default");
        if (p.Title is { } t) sb.Append($".WithTitle({RenderString(t)})");
        if (p.Author is { } a) sb.Append($".WithAuthor({RenderString(a)})");
        if (p.Subject is { } s) sb.Append($".WithSubject({RenderString(s)})");
        if (p.Keywords is { } k) sb.Append($".WithKeywords({RenderString(k)})");
        if (p.Comments is { } c) sb.Append($".WithComments({RenderString(c)})");
        if (p.Category is { } cat) sb.Append($".WithCategory({RenderString(cat)})");
        if (p.Company is { } co) sb.Append($".WithCompany({RenderString(co)})");
        return sb.ToString();
    }

    // ----- Top level ------------------------------------------------------------------------

    /// <summary>Renders <paramref name="document"/> as a self-contained file-based-app C#
    /// file - see <see cref="DocumentIO"/> and this class's own doc comment.</summary>
    public static string Generate(IEnumerable<string> referenceLines, string outputFileName, Document document)
    {
        var sb = new StringBuilder();

        foreach (var line in referenceLines)
            sb.AppendLine(line);

        sb.AppendLine();
        sb.AppendLine("using System;");
        sb.AppendLine("using Kookerella.CsWordDsl;");
        sb.AppendLine();

        sb.Append("var document = Document.Create(").Append(string.Join(", ", document.Sections.Select(RenderSection))).AppendLine(")");

        if (document.Styles.Count > 0 && !document.Styles.SequenceEqual(BuiltInStyles.All))
            sb.Append("    .WithStyles(").Append(string.Join(", ", document.Styles.Select(RenderStyleDefinition))).AppendLine(")");

        if (document.Numbering.Count > 0)
            sb.Append("    .WithNumbering(").Append(string.Join(", ", document.Numbering.Select(RenderNumberingDefinition))).AppendLine(")");

        if (document.Protection is { } protection)
            sb.Append("    .WithProtection(").Append(RenderDocumentProtection(protection)).AppendLine(")");

        if (document.VbaProject is { } vba)
            sb.Append("    .WithVbaProject(").Append(RenderByteArray(vba)).AppendLine(")");

        if (document.Properties != DocumentProperties.Default)
            sb.Append("    .WithDocumentProperties(").Append(RenderDocumentProperties(document.Properties)).AppendLine(")");

        if (document.TableStyles.Count > 0)
            sb.Append("    .WithTableStyles(").Append(string.Join(", ", document.TableStyles.Select(RenderTableStyleDefinition))).AppendLine(")");

        sb.AppendLine("    ;");
        sb.AppendLine();
        sb.AppendLine($"DocumentIO.Save(document, {RenderString(outputFileName)});");

        return sb.ToString();
    }
}
