using Xunit;
using static Kookerella.CsWordDsl.Tests.TestHelpers;

namespace Kookerella.CsWordDsl.Tests;

/// <summary>Behavioral round-trip tests against throwaway temp files - each saves a
/// <see cref="Document"/>, asserts the file is schema-valid, reloads it, and checks the
/// specific values that scenario cares about. Deliberately targeted assertions rather than
/// whole-<see cref="Document"/> equality: this wrapper's records use <c>IReadOnlyList&lt;T&gt;</c>
/// properties, whose compiler-synthesized record equality is reference-based for the list
/// itself (not a deep structural comparison), the same limitation the sibling
/// Kookerella.CsOpenXmlDsl.Tests project's own <c>WorkbookTests</c> works around the same
/// way.</summary>
public class DocumentTests
{
    [Fact]
    public void Basic_paragraphs_and_runs_round_trip()
    {
        var doc = Document.Create(
            Section.Of([
                Block.Paragraph([new Inline.Run("Title Here")], styleId: "Title"),
                Block.Paragraph([
                    new Inline.Run("Plain, "),
                    new Inline.Run("bold, ", new RunStyle { Bold = true }),
                    new Inline.Run("italic.", new RunStyle { Italic = true })
                ])
            ]));

        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        Assert.Single(loaded.Sections);
        var body = loaded.Sections[0].Body;
        Assert.Equal(2, body.Count);

        var titlePara = ((Block.ParagraphBlock)body[0]).Para;
        Assert.Equal("Title", titlePara.StyleId);
        Assert.Equal("Title Here", ((Inline.Run)titlePara.Inlines[0]).Text);

        var para2 = ((Block.ParagraphBlock)body[1]).Para;
        Assert.True(((Inline.Run)para2.Inlines[1]).Style?.Bold);
        Assert.True(((Inline.Run)para2.Inlines[2]).Style?.Italic);
    }

    [Fact]
    public void Hyperlink_round_trips()
    {
        var doc = Document.Create(Section.Of([Block.Paragraph([Inline.HyperlinkText("click here", new HyperlinkTarget.ExternalUrl("https://example.com/data"))])]));
        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var para = ((Block.ParagraphBlock)loaded.Sections[0].Body[0]).Para;
        var link = (Inline.Hyperlink)para.Inlines[0];
        Assert.Equal("https://example.com/data", ((HyperlinkTarget.ExternalUrl)link.Target).Url);
        Assert.Equal("click here", ((Inline.Run)link.Runs[0]).Text);
    }

    [Fact]
    public void Bookmark_round_trips()
    {
        var doc = Document.Create(Section.Of([Block.Paragraph([new Inline.Bookmark("Section1", [new Inline.Run("bookmarked text")])])]));
        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var bm = (Inline.Bookmark)((Block.ParagraphBlock)loaded.Sections[0].Body[0]).Para.Inlines[0];
        Assert.Equal("Section1", bm.Name);
        Assert.Equal("bookmarked text", ((Inline.Run)bm.Content[0]).Text);
    }

    [Fact]
    public void Comment_round_trips()
    {
        var doc = Document.Create(Section.Of([Block.Paragraph([new Inline.Comment("Alex", "AR", null, "Please review.", [new Inline.Run("flagged text")])])]));
        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var comment = (Inline.Comment)((Block.ParagraphBlock)loaded.Sections[0].Body[0]).Para.Inlines[0];
        Assert.Equal("Alex", comment.Author);
        Assert.Equal("AR", comment.Initials);
        Assert.Equal("Please review.", comment.Text);
        Assert.Equal("flagged text", ((Inline.Run)comment.Content[0]).Text);
    }

    [Fact]
    public void Table_with_merges_and_style_round_trips()
    {
        var table = Block.Table(
            [
                TableRow.Of([TableCell.Of([Block.Paragraph([new Inline.Run("Header")])], new TableCellProps { GridSpan = 2 })], repeatAsHeader: true),
                TableRow.Of(
                [
                    TableCell.Of([Block.Paragraph([new Inline.Run("A")])], new TableCellProps { VerticalMerge = VerticalMergeKind.Restart }),
                    TableCell.Of([Block.Paragraph([new Inline.Run("B")])])
                ])
            ],
            [150.0, 150.0],
            style: TableStyleRef.Named("TableGrid"));

        var doc = Document.Create(Section.Of([table]));
        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var loadedTable = ((Block.TableBlock)loaded.Sections[0].Body[0]).Entry;
        Assert.Equal(2, loadedTable.Rows.Count);
        Assert.True(loadedTable.Rows[0].RepeatAsHeader);
        Assert.Equal(2, loadedTable.Rows[0].Cells[0].Props.GridSpan);
        Assert.Equal(VerticalMergeKind.Restart, loadedTable.Rows[1].Cells[0].Props.VerticalMerge);
        Assert.Equal("TableGrid", loadedTable.Style?.Name);
    }

    [Fact]
    public void Image_round_trips()
    {
        var image = ImageEntry.FromBytesInches(OnePixelPng(), ImageFormat.Png, 1.0, 1.0, "a pixel");
        var doc = Document.Create(Section.Of([Block.Paragraph([new Inline.Image(image)])]));
        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var img = (Inline.Image)((Block.ParagraphBlock)loaded.Sections[0].Body[0]).Para.Inlines[0];
        Assert.Equal(OnePixelPng(), img.Entry.Data);
        Assert.Equal("a pixel", img.Entry.AltText);
    }

    [Fact]
    public void Footnote_and_endnote_round_trip()
    {
        var doc = Document.Create(Section.Of([Block.Paragraph([new Inline.Run("See note."), Inline.FootnoteText("A footnote."), Inline.EndnoteText("An endnote.")])]));
        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var inlines = ((Block.ParagraphBlock)loaded.Sections[0].Body[0]).Para.Inlines;
        var footnote = (Inline.Footnote)inlines[1];
        var endnote = (Inline.Endnote)inlines[2];
        Assert.Equal("A footnote.", ((Inline.Run)((Block.ParagraphBlock)footnote.Content[0]).Para.Inlines[0]).Text);
        Assert.Equal("An endnote.", ((Inline.Run)((Block.ParagraphBlock)endnote.Content[0]).Para.Inlines[0]).Text);
    }

    [Fact]
    public void Track_changes_round_trip()
    {
        var editDate = new DateTime(2024, 3, 1, 14, 0, 0);
        var doc = Document.Create(Section.Of([
            Block.Paragraph([
                new Inline.Run("The "),
                Inline.Inserted([new Inline.Run("quick ")], "Alex", editDate),
                Inline.Deleted([new Inline.Run("slow ")], "Alex", editDate),
                new Inline.Run("fox.")
            ])
        ]));

        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var inlines = ((Block.ParagraphBlock)loaded.Sections[0].Body[0]).Para.Inlines;
        var inserted = (Inline.TrackedChange)inlines[1];
        var deleted = (Inline.TrackedChange)inlines[2];
        Assert.Equal(RevisionKind.Inserted, inserted.Revision.Kind);
        Assert.Equal("Alex", inserted.Revision.Author);
        Assert.Equal(RevisionKind.Deleted, deleted.Revision.Kind);
        Assert.Equal("quick ", ((Inline.Run)inserted.Content[0]).Text);
    }

    [Fact]
    public void Content_controls_round_trip()
    {
        var doc = Document.Create(Section.Of([
            Block.Paragraph([
                new Inline.ContentControl(
                    new ContentControlProps { Alias = "Name", Tag = "name", Type = new ContentControlType.PlainText(false), Lock = ContentControlLock.LockContentEditing },
                    [new Inline.Run("Type here")]),
                new Inline.ContentControl(
                    new ContentControlProps { Type = new ContentControlType.DropDown([("Yes", "yes"), ("No", "no")], Editable: true) },
                    [new Inline.Run("Yes")]),
                new Inline.ContentControl(
                    new ContentControlProps { Type = new ContentControlType.CheckBox(true, ("Wingdings", "2612"), ("Wingdings", "2610")) },
                    [new Inline.Run("☒")])
            ])
        ]));

        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var inlines = ((Block.ParagraphBlock)loaded.Sections[0].Body[0]).Para.Inlines;

        var plainText = (Inline.ContentControl)inlines[0];
        Assert.Equal("Name", plainText.Props.Alias);
        Assert.Equal(ContentControlLock.LockContentEditing, plainText.Props.Lock);
        Assert.IsType<ContentControlType.PlainText>(plainText.Props.Type);

        var dropDown = (ContentControlType.DropDown)((Inline.ContentControl)inlines[1]).Props.Type;
        Assert.True(dropDown.Editable);
        Assert.Equal(2, dropDown.Items.Count);

        var checkBox = (ContentControlType.CheckBox)((Inline.ContentControl)inlines[2]).Props.Type;
        Assert.True(checkBox.Checked);
        Assert.Equal(("Wingdings", "2612"), checkBox.CheckedSymbol);
    }

    [Fact]
    public void Page_setup_and_sections_round_trip()
    {
        var doc = Document.Create(
            Section.With(
                SectionProperties.Default.WithOrientation(PageOrientation.Landscape).WithPageSize(new PageSize.A4()).WithColumns(2),
                [Block.Paragraph([new Inline.Run("Landscape A4, two columns.")])]));

        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var props = loaded.Sections[0].Properties;
        Assert.Equal(PageOrientation.Landscape, props.Orientation);
        Assert.IsType<PageSize.A4>(props.PageSize);
        Assert.Equal(2, props.Columns);
    }

    [Fact]
    public void Document_protection_round_trips()
    {
        var doc = Document.Create(Section.Of([Block.Paragraph([new Inline.Run("Protected.")])]))
            .WithProtection(DocumentProtection.With(EditRestriction.ReadOnly, "hunter2"));

        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        // Password never round-trips (the hash isn't reversible) - see the F# core's own
        // DocumentProtection.Password doc comment.
        Assert.Equal(EditRestriction.ReadOnly, loaded.Protection?.Edit);
    }

    [Fact]
    public void Vba_project_bytes_round_trip()
    {
        var macroBytes = new byte[] { 1, 2, 3, 4, 5 };
        var doc = Document.Create(Section.Of([Block.Paragraph([new Inline.Run("Has a macro.")])])).WithVbaProject(macroBytes);
        var path = Path.Combine(Path.GetTempPath(), $"CsWordDslTest_{Guid.NewGuid():N}.docm");
        DocumentIO.Save(doc, path);
        var loaded = DocumentIO.Load(path);

        Assert.Equal(macroBytes, loaded.VbaProject);
    }

    [Fact]
    public void Document_properties_round_trip()
    {
        var doc = Document.Create(Section.Of([Block.Paragraph([new Inline.Run("x")])]))
            .WithDocumentProperties(DocumentProperties.Default.WithTitle("Quarterly Report").WithAuthor("Kookerella"));

        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        Assert.Equal("Quarterly Report", loaded.Properties.Title);
        Assert.Equal("Kookerella", loaded.Properties.Author);
    }

    [Fact]
    public void Numbered_list_round_trips()
    {
        var listDef = NumberingDefinition.NumberedList(1);
        var doc = Document.Create(Section.Of([
            Block.Paragraph([new Inline.Run("First")], numbering: (1, 0)),
            Block.Paragraph([new Inline.Run("Second")], numbering: (1, 0))
        ])).WithNumbering(listDef);

        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        Assert.Single(loaded.Numbering);
        Assert.Equal((1, 0), ((Block.ParagraphBlock)loaded.Sections[0].Body[0]).Para.Numbering);
    }

    [Fact]
    public void Table_style_definition_round_trips()
    {
        var customStyle = new TableStyleDefinition { Id = "MyStyle", Name = "My Style" }
            .WithBandedRow(TableStyleRegion.None.WithCellShading(Color.Black))
            .WithBandedRow2(TableStyleRegion.None.WithCellShading(Color.White));

        var doc = Document.Create(Section.Of([
            Block.Table([TableRow.Of([TableCell.Of([Block.Paragraph([new Inline.Run("x")])])])], [100.0], style: TableStyleRef.Named("MyStyle"))
        ])).WithTableStyles(customStyle);

        var path = TempDocxPath();
        DocumentIO.Save(doc, path);
        AssertSchemaValid(path);
        var loaded = DocumentIO.Load(path);

        var loadedStyle = Assert.Single(loaded.TableStyles);
        Assert.Equal("MyStyle", loadedStyle.Id);
        Assert.NotNull(loadedStyle.BandedRow.CellShading);
        Assert.NotNull(loadedStyle.BandedRow2.CellShading);
    }
}
