module Kookerella.FsWordDsl.Tests

open System
open System.Diagnostics
open System.IO
open System.Xml.Linq
open Xunit
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Validation
open Kookerella.FsWordDsl
open XmlTests
open JsonTests
open type Kookerella.FsWordDsl.DocumentDsl

// --- Scenario harness -------------------------------------------------------------
//
// Each scenario below is a self-contained demonstration of one feature. Running it
// writes the document it builds to Examples/<scenario name>/output.docx (checked into
// the repo) so you can open any single feature in Word without re-running anything,
// while the test itself verifies the file is schema-valid and round-trips exactly back
// through the DSL - the same shape Excel's own `Tests.fs` harness uses.

let private examplesDir = Path.Combine(__SOURCE_DIRECTORY__, "Examples")

let private assertSchemaValid (path: string) =
    use doc = WordprocessingDocument.Open(path, false)
    // Office2010, not the parameterless (Office2007) default - `w:tblLook`'s named
    // boolean attributes (firstRow/lastRow/noHBand/noVBand), which real modern Word
    // itself writes, aren't valid under the older Office2007 transitional schema.
    let validator = OpenXmlValidator(FileFormatVersions.Office2010)
    let errors = validator.Validate(doc) |> List.ofSeq

    Assert.True(
        errors.IsEmpty,
        String.Join("\n", errors |> Seq.map (fun e -> sprintf "%s: %s" e.Path.XPath e.Description))
    )

/// `CommentRangeStart`/`CommentRangeEnd`'s own `id` is documented as write-time-only (see
/// `Model.Inline.CommentRangeStart`'s own doc comment) - `Reader` reconstructs some id from
/// the real OOXML `w:id`, which generally won't match what a caller originally wrote, so
/// `verifyScenarioNamed`'s own round-trip assertion normalizes it away on both sides first,
/// the same "known-lossy field, compare with it blanked out" treatment `Password` gets.
let private normalizeCommentRangeIds (sections: Section list) : Section list =
    let rec normalizeInline (i: Inline) : Inline =
        match i with
        | Hyperlink(target, runs, tooltip) -> Hyperlink(target, runs |> List.map normalizeInline, tooltip)
        | Bookmark(name, content) -> Bookmark(name, content |> List.map normalizeInline)
        | Comment(author, initials, date, text, content) -> Comment(author, initials, date, text, content |> List.map normalizeInline)
        | CommentRangeStart(_, author, initials, date, text) -> CommentRangeStart("_", author, initials, date, text)
        | CommentRangeEnd _ -> CommentRangeEnd "_"
        | Footnote content -> Footnote(content |> List.map normalizeBlock)
        | Endnote content -> Endnote(content |> List.map normalizeBlock)
        | other -> other

    and normalizeBlock (b: Block) : Block =
        match b with
        | ParagraphBlock p -> ParagraphBlock { p with Inlines = p.Inlines |> List.map normalizeInline }
        | TableBlock t ->
            TableBlock
                { t with
                    Rows =
                        t.Rows
                        |> List.map (fun r ->
                            { r with Cells = r.Cells |> List.map (fun c -> { c with Content = c.Content |> List.map normalizeBlock }) }) }

    let normalizeHeaderFooterSet (h: HeaderFooterSet) : HeaderFooterSet =
        { Default = h.Default |> Option.map (List.map normalizeBlock)
          First = h.First |> Option.map (List.map normalizeBlock)
          Even = h.Even |> Option.map (List.map normalizeBlock) }

    sections
    |> List.map (fun s ->
        { s with
            Body = s.Body |> List.map normalizeBlock
            Properties =
                { s.Properties with
                    Header = s.Properties.Header |> Option.map normalizeHeaderFooterSet
                    Footer = s.Properties.Footer |> Option.map normalizeHeaderFooterSet } })

let private codeGenReferenceLines =
    [ sprintf "#r \"%s\"" (typeof<Document>.Assembly.Location.Replace("\\", "\\\\"))
      sprintf "#r \"%s\"" (typeof<WordprocessingDocument>.Assembly.Location.Replace("\\", "\\\\")) ]

/// Saves `doc` to `Examples/<name>/<fileName>`, asserts the file is schema-valid, and
/// asserts it round-trips exactly back through the DSL. Also writes `Examples/<name>/
/// script.fsx` (see the `Category=Slow` tests below), `document.xml`, and
/// `document.json` - one folder always has four views of the same example.
let private verifyScenarioNamed (name: string) (fileName: string) (doc: Document) =
    let dir = Path.Combine(examplesDir, name)
    Directory.CreateDirectory(dir) |> ignore
    let path = Path.Combine(dir, fileName)
    Document.save path doc

    assertSchemaValid path

    let roundTripped = Document.load path
    Assert.Equal<Section list>(doc.Sections |> normalizeCommentRangeIds, roundTripped.Sections |> normalizeCommentRangeIds)
    Assert.Equal<StyleDefinition list>(doc.Styles, roundTripped.Styles)
    Assert.Equal<NumberingDefinition list>(doc.Numbering, roundTripped.Numbering)

    // Password never round-trips (the hash isn't reversible - see DocumentProtection's
    // own doc comment), so compare with it normalized away on both sides.
    Assert.Equal<DocumentProtection option>(
        doc.Protection |> Option.map (fun p -> { p with Password = None }),
        roundTripped.Protection
    )

    // F#'s structural equality on `option`/array values compares by content, not
    // reference, so this is a genuine byte-for-byte comparison of the VBA project.
    Assert.Equal<byte[] option>(doc.VbaProject, roundTripped.VbaProject)

    Assert.Equal<DocumentProperties>(doc.Properties, roundTripped.Properties)
    Assert.Equal<TableStyleDefinition list>(doc.TableStyles, roundTripped.TableStyles)

    let script = Document.generateScript codeGenReferenceLines fileName doc
    File.WriteAllText(Path.Combine(dir, "script.fsx"), script)

    let xml = Xml.toDocument doc
    assertXmlSchemaValid (XDocument(xml))
    xml.Save(Path.Combine(dir, "document.xml"))

    let json = Json.toDocument doc
    assertJsonSchemaValid json
    File.WriteAllText(Path.Combine(dir, "document.json"), json.ToJsonString())

// A minimal, valid 1x1 transparent PNG - real enough for a real embedded image part,
// tiny enough to inline here rather than a separate binary asset.
let private onePixelPng =
    Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
    )

// --- Scenarios ----------------------------------------------------------------------

[<Fact>]
let ``BasicParagraphsAndRuns`` () =
    let doc =
        document
            [ section
                  [ para
                        [ run "Plain text, "
                          run ("bold", style = { RunStyle.Default with Bold = true })
                          run ", "
                          run ("italic", style = { RunStyle.Default with Italic = true })
                          run ", and "
                          run ("colored", style = { RunStyle.Default with Color = Some Color.red }) ]
                    para
                        [ run ("Underlined", style = { RunStyle.Default with Underline = Some SingleUnderline })
                          run " and "
                          run ("struck through", style = { RunStyle.Default with Strikethrough = true }) ]
                    para
                        [ run ("SMALL CAPS", style = { RunStyle.Default with SmallCaps = true })
                          run " and "
                          run ("ALL CAPS", style = { RunStyle.Default with AllCaps = true })
                          run " and "
                          run ("hidden text", style = { RunStyle.Default with Hidden = true }) ]
                    para
                        [ run (
                              "Theme-colored text (Accent1, tinted 20%)",
                              style = { RunStyle.Default with Color = Some(Theme(Accent1Theme, (0x1Fuy, 0x49uy, 0x7Duy), Some 0.2, None)) }
                          ) ] ] ]

    verifyScenarioNamed "BasicParagraphsAndRuns" "output.docx" doc

[<Fact>]
let ``NamedStyles`` () =
    let doc =
        document
            [ section
                  [ para ([ run "Document Title" ], styleId = "Title")
                    para ([ run "Section Heading" ], styleId = "Heading1")
                    para [ run "This paragraph uses the default Normal style." ] ] ]

    verifyScenarioNamed "NamedStyles" "output.docx" doc

[<Fact>]
let ``ParagraphFormatting`` () =
    let centered =
        { ParagraphFormat.Default with
            Alignment = Some AlignCenter
            SpacingBefore = Some 12.0
            SpacingAfter = Some 12.0 }

    let indented =
        { ParagraphFormat.Default with
            Indentation = Some { Indentation.None with Left = Some 36.0 }
            LineSpacing = Some DoubleSpacing }

    let bordered =
        { ParagraphFormat.Default with
            Borders =
                Some
                    { BorderStyle.None with
                        Top = Some { Style = SingleLine; Width = Some 1.0; Color = Some Color.black }
                        Bottom = Some { Style = SingleLine; Width = Some 1.0; Color = Some Color.black } }
            Shading = Some(Rgb(0xD9uy, 0xD9uy, 0xD9uy)) }

    let tabbed =
        { ParagraphFormat.Default with
            TabStops =
                [ { Position = 288.0; Alignment = RightTab; Leader = DotLeader } ] }

    let themeShaded =
        { ParagraphFormat.Default with
            Shading = Some(Theme(Accent6Theme, (0xF2uy, 0xDCuy, 0xDBuy), None, Some 0.4)) }

    let doc =
        document
            [ section
                  [ para ([ run "Centered heading-like text." ], format = centered)
                    para ([ run "An indented, double-spaced paragraph." ], format = indented)
                    para ([ run "A paragraph with top/bottom borders and shading." ], format = bordered)
                    para ([ run "Introduction"; Tab; run "1" ], format = tabbed)
                    para ([ run "A paragraph shaded with a theme color (Accent6, shaded 40%)." ], format = themeShaded) ] ]

    verifyScenarioNamed "ParagraphFormatting" "output.docx" doc

[<Fact>]
let ``BulletList`` () =
    let doc =
        document
            [ section
                  [ para ([ run "First bullet" ], numbering = (1, 0))
                    para ([ run "Second bullet" ], numbering = (1, 0))
                    para ([ run "Third bullet" ], numbering = (1, 0)) ] ]
        |> withNumbering [ bulletListDef 1 ]

    verifyScenarioNamed "BulletList" "output.docx" doc

[<Fact>]
let ``NumberedList`` () =
    let doc =
        document
            [ section
                  [ para ([ run "First step" ], numbering = (1, 0))
                    para ([ run "Second step" ], numbering = (1, 0))
                    para ([ run "Third step" ], numbering = (1, 0)) ] ]
        |> withNumbering [ numberedListDef 1 ]

    verifyScenarioNamed "NumberedList" "output.docx" doc

[<Fact>]
let ``MultiLevelNumberedList`` () =
    let doc =
        document
            [ section
                  [ para ([ run "First topic" ], numbering = (1, 0))
                    para ([ run "First subtopic" ], numbering = (1, 1))
                    para ([ run "First sub-subtopic" ], numbering = (1, 2))
                    para ([ run "Second subtopic" ], numbering = (1, 1))
                    para ([ run "Second topic" ], numbering = (1, 0)) ] ]
        |> withNumbering [ multiLevelNumberedListDef 1 3 ]

    verifyScenarioNamed "MultiLevelNumberedList" "output.docx" doc

[<Fact>]
let ``Table_Basic`` () =
    // Every real .docx cell carries an explicit width (`w:tcW`) whether or not the DSL
    // caller set one - `Props.Width = None` still resolves to the column's own width at
    // write time (see `tableCellToW`'s `colWidthTwips` fallback), so it's given explicitly
    // here to round-trip exactly rather than relying on that fallback.
    let headerCell (width: float) (text: string) =
        tableCell (
            [ para [ run (text, style = { RunStyle.Default with Bold = true }) ] ],
            props = { TableCellProps.Default with Shading = Some(Rgb(220uy, 220uy, 220uy)); Width = Some width }
        )

    let bodyCell (width: float) (text: string) =
        tableCell ([ para [ run text ] ], props = { TableCellProps.Default with Width = Some width })

    let doc =
        document
            [ section
                  [ table (
                        [ tableRow [ headerCell 200.0 "Item"; headerCell 100.0 "Quantity" ]
                          tableRow [ bodyCell 200.0 "Widgets"; bodyCell 100.0 "12" ]
                          tableRow [ bodyCell 200.0 "Gadgets"; bodyCell 100.0 "5" ] ],
                        [ 200.0; 100.0 ],
                        style = TableStyleRef.Default
                    ) ] ]

    verifyScenarioNamed "Table_Basic" "output.docx" doc

[<Fact>]
let ``Table_MergedCells`` () =
    let cell (content: Block list) (props: TableCellProps) = tableCell (content, props = { props with Width = Some 150.0 })

    let doc =
        document
            [ section
                  [ table (
                        [ tableRow
                              [ cell
                                    [ para [ run ("Spans 2 columns", style = { RunStyle.Default with Bold = true }) ] ]
                                    { TableCellProps.Default with GridSpan = Some 2 } ]
                          tableRow
                              [ cell [ para [ run "A1" ] ] TableCellProps.Default
                                cell [ para [ run "B1 (merged down)" ] ] { TableCellProps.Default with VerticalMerge = Some RestartMerge } ]
                          tableRow
                              [ cell [ para [ run "A2" ] ] TableCellProps.Default
                                cell [ para [] ] { TableCellProps.Default with VerticalMerge = Some ContinueMerge } ] ],
                        [ 150.0; 150.0 ]
                    ) ] ]

    verifyScenarioNamed "Table_MergedCells" "output.docx" doc

[<Fact>]
let ``Table_BordersAndStyle`` () =
    let thick: BorderSide = { Style = ThickLine; Width = Some 2.0; Color = Some Color.black }
    let thin: BorderSide = { thick with Width = Some 0.5 }

    let borders: TableBorders =
        { Outer =
            { Left = Some thick
              Right = Some thick
              Top = Some thick
              Bottom = Some thick }
          InsideHorizontal = Some thin
          InsideVertical = Some thin }

    let cell (text: string) = tableCell ([ para [ run text ] ], props = { TableCellProps.Default with Width = Some 150.0 })

    let doc =
        document
            [ section
                  [ table (
                        [ tableRow [ cell "A"; cell "B" ]
                          tableRow [ cell "C"; cell "D" ] ],
                        [ 150.0; 150.0 ],
                        style = { TableStyleRef.Default with Name = "TableGrid"; BandedRows = true },
                        borders = borders
                    ) ] ]

    verifyScenarioNamed "Table_BordersAndStyle" "output.docx" doc

[<Fact>]
let ``Image`` () =
    let img: ImageEntry =
        { Data = onePixelPng
          Format = Png
          WidthEmu = Units.inchesToEmu 1.0
          HeightEmu = Units.inchesToEmu 1.0
          AltText = Some "A single red pixel, scaled up" }

    let doc = document [ section [ para [ run "Here is an image: "; image img ] ] ]
    verifyScenarioNamed "Image" "output.docx" doc

[<Fact>]
let ``Hyperlink_External`` () =
    let doc =
        document
            [ section [ para [ run "Visit "; hyperlink ("Kookerella on GitHub", ExternalUrl "https://github.com/Kookerella-Ltd"); run " for more." ] ] ]

    verifyScenarioNamed "Hyperlink_External" "output.docx" doc

[<Fact>]
let ``Bookmark`` () =
    let doc = document [ section [ para [ bookmark ("TopOfDocument", [ run "This paragraph is bookmarked." ]) ] ] ]
    verifyScenarioNamed "Bookmark" "output.docx" doc

[<Fact>]
let ``Bookmark_MultiParagraph`` () =
    let doc =
        document
            [ section
                  [ para [ run "Before the bookmark." ]
                    para [ BookmarkRangeStart "Section2"; run "The bookmark starts on this paragraph" ]
                    para [ run "and continues through this one" ]
                    para [ run "and ends on this one."; BookmarkRangeEnd "Section2" ]
                    para [ run "After the bookmark." ] ] ]

    verifyScenarioNamed "Bookmark_MultiParagraph" "output.docx" doc

[<Fact>]
let ``Hyperlink_Internal`` () =
    let doc =
        document
            [ section
                  [ para [ hyperlink ("Jump to the target below", InternalBookmark "Target") ]
                    para [ bookmark ("Target", [ run "You made it." ]) ] ] ]

    verifyScenarioNamed "Hyperlink_Internal" "output.docx" doc

[<Fact>]
let ``Comments`` () =
    let doc =
        document
            [ section
                  [ para
                        [ comment (
                              [ run "This figure needs review." ],
                              "Please double check the totals.",
                              author = "Alex",
                              initials = "AR",
                              // Explicit, rather than relying on the "None -> now at write
                              // time" default (see `Model.Inline.Comment`'s own doc
                              // comment) - the point of this test is round-trip fidelity,
                              // and "now" is different on every run by construction.
                              date = DateTime(2024, 1, 15, 9, 30, 0)
                          ) ] ] ]

    verifyScenarioNamed "Comments" "output.docx" doc

[<Fact>]
let ``Comments_MultiParagraph`` () =
    let doc =
        document
            [ section
                  [ para [ run "Before the comment." ]
                    para
                        [ CommentRangeStart("review1", "Alex", Some "AR", Some(DateTime(2024, 1, 15, 9, 30, 0)), "This whole section needs a second look.")
                          run "The comment starts on this paragraph" ]
                    para [ run "and continues through this one" ]
                    para [ run "and ends on this one."; CommentRangeEnd "review1" ]
                    para [ run "After the comment." ] ] ]

    verifyScenarioNamed "Comments_MultiParagraph" "output.docx" doc

[<Fact>]
let ``TrackedChanges`` () =
    let editDate = DateTime(2024, 3, 1, 14, 0, 0)

    let doc =
        document
            [ section
                  [ para
                        [ run "The quick "
                          inserted ([ run "brown " ], "Alex", editDate)
                          run "fox jumps over the "
                          deleted ([ run "lazy " ], "Alex", editDate)
                          run "dog." ]
                    para
                        (
                            [ run "This whole paragraph was inserted." ],
                            markRevision = { Kind = Inserted; Author = "Alex"; Date = Some editDate }
                        ) ] ]

    verifyScenarioNamed "TrackedChanges" "output.docx" doc

[<Fact>]
let ``FootnotesAndEndnotes`` () =
    let props =
        { SectionProperties.Default with
            FootnoteNumbering = Some { Format = LowerRomanFormat; StartAt = None; Restart = RestartEachPage }
            EndnoteNumbering = Some { Format = DecimalFormat; StartAt = Some 100; Restart = ContinuousRestart } }

    let doc =
        document
            [ sectionWith
                  props
                  [ para
                        [ run "This claim needs a citation"
                          footnote "Smith, J. (2023). A Study of Claims. Journal of Claims, 12(3), 45-67."
                          run ", and this one refers to a fuller discussion"
                          endnote [ para [ run "See the appendix for the full derivation and worked examples." ] ] ]
                    para [ run "A second footnote in a later paragraph, to check id numbering isn't reset per paragraph."; footnote "A second note." ] ] ]

    verifyScenarioNamed "FootnotesAndEndnotes" "output.docx" doc

[<Fact>]
let ``PageSetupLandscape`` () =
    let props =
        { SectionProperties.Default with
            Orientation = Landscape
            Margins = { PageMargins.Default with Left = 54.0; Right = 54.0 } }

    let doc = document [ sectionWith props [ para [ run "A landscape-oriented page." ] ] ]
    verifyScenarioNamed "PageSetupLandscape" "output.docx" doc

[<Fact>]
let ``HeaderFooterDefault`` () =
    let footer: HeaderFooterSet =
        { HeaderFooterSet.None with Default = Some [ para [ run "Page "; Field("PAGE", Some "1") ] ] }

    let props = { SectionProperties.Default with Footer = Some footer }
    let doc = document [ sectionWith props [ para [ run "Body text with a page-number footer." ] ] ]
    verifyScenarioNamed "HeaderFooterDefault" "output.docx" doc

[<Fact>]
let ``HeaderFooterFirstPageDifferent`` () =
    let header: HeaderFooterSet =
        { HeaderFooterSet.None with
            Default = Some [ para [ run "Standard Header" ] ]
            First = Some [ para [ run "Cover Page" ] ] }

    let props = { SectionProperties.Default with Header = Some header }

    let doc =
        document
            [ sectionWith
                  props
                  [ para [ run "First-page content."; PageBreak; run "Content after a manual page break." ] ] ]

    verifyScenarioNamed "HeaderFooterFirstPageDifferent" "output.docx" doc

[<Fact>]
let ``MultipleSections`` () =
    // `BreakType` describes how a section begins relative to the PREVIOUS one, so it's
    // meaningless (and not written) on the very first section - sec2/sec3 each show a
    // different real break type.
    let sec1 = sectionWith { SectionProperties.Default with Orientation = Portrait } [ para [ run "Section 1 - portrait." ] ]

    let sec2 =
        sectionWith
            { SectionProperties.Default with Orientation = Landscape; BreakType = ContinuousBreak }
            [ para [ run "Section 2 - landscape, continuous break (no page break from section 1)." ] ]

    let sec3 =
        sectionWith
            { SectionProperties.Default with Orientation = Portrait; BreakType = OddPageBreak }
            [ para [ run "Section 3 - portrait again, starts on the next odd page." ] ]

    let doc = document [ sec1; sec2; sec3 ]
    verifyScenarioNamed "MultipleSections" "output.docx" doc

[<Fact>]
let ``DocumentProtectionReadOnly`` () =
    let doc =
        document [ section [ para [ run "This document is protected as read-only." ] ] ]
        |> withProtection { Edit = Some ReadOnlyRestriction; Password = Some "hunter2" }

    verifyScenarioNamed "DocumentProtectionReadOnly" "output.docm" doc

[<Fact>]
let ``Macro`` () =
    // Synthetic bytes, not a real compiled VBA project - Core embeds and reads back
    // whatever bytes it's given verbatim (see `Document.VbaProject`'s own doc comment),
    // so this still exercises the real round trip; unlike Excel's own `VbaMacro`
    // scenario, no real Word-produced `vbaProject.bin` test asset was available in this
    // environment to substitute for full realism.
    let doc =
        document [ section [ para [ run "This is a macro-enabled template placeholder." ] ] ]
        |> withVbaProject [| 1uy; 2uy; 3uy; 4uy; 5uy |]

    verifyScenarioNamed "Macro" "output.docm" doc

[<Fact>]
let ``DocumentProperties`` () =
    let doc =
        document [ section [ para [ run "A document with core metadata set." ] ] ]
        |> withDocumentProperties
            { Title = Some "Quarterly Report"
              Author = Some "Kookerella"
              Subject = Some "Q3 Results"
              Keywords = Some "finance, quarterly, report"
              Comments = Some "Draft for review"
              Category = Some "Reports"
              Company = Some "Kookerella Ltd" }

    verifyScenarioNamed "DocumentProperties" "output.docx" doc

[<Fact>]
let ``Table_CustomStyleAndHeaderRow`` () =
    let thin: BorderSide = { Style = SingleLine; Width = Some 0.5; Color = Some Color.black }

    let customStyle: TableStyleDefinition =
        { TableStyleDefinition.Default with
            Id = "MyTableStyle"
            Name = "My Table Style"
            Borders =
                Some
                    { Outer = { Left = Some thin; Right = Some thin; Top = Some thin; Bottom = Some thin }
                      InsideHorizontal = Some thin
                      InsideVertical = Some thin }
            FirstRow =
                { TableStyleRegion.None with
                    RunFormat = Some { RunStyle.Default with Bold = true; Color = Some Color.white }
                    CellShading = Some(Rgb(0x4Fuy, 0x81uy, 0xBDuy)) }
            LastRow = { TableStyleRegion.None with RunFormat = Some { RunStyle.Default with Italic = true } }
            FirstColumn = { TableStyleRegion.None with RunFormat = Some { RunStyle.Default with Bold = true } }
            BandedRow = { TableStyleRegion.None with CellShading = Some(Rgb(0xDCuy, 0xE6uy, 0xF1uy)) }
            BandedColumn = { TableStyleRegion.None with CellShading = Some(Rgb(0xF2uy, 0xF2uy, 0xF2uy)) }
            NorthWestCell = { TableStyleRegion.None with CellShading = Some(Rgb(0x2Duy, 0x50uy, 0x82uy)) } }

    let cell (text: string) (width: float) = tableCell ([ para [ run text ] ], props = { TableCellProps.Default with Width = Some width })

    let amountCell =
        tableCell (
            [ para [ run "Amount" ] ],
            props = { TableCellProps.Default with Width = Some 150.0; Margins = Some { Top = Some 2.0; Bottom = Some 2.0; Left = Some 8.0; Right = Some 8.0 } }
        )

    let doc =
        document
            [ section
                  [ table (
                        [ tableRow ([ cell "Name" 200.0; amountCell ], height = 20.0, repeatAsHeader = true)
                          tableRow [ cell "Widgets" 200.0; cell "12" 150.0 ]
                          tableRow [ cell "Gadgets" 200.0; cell "7" 150.0 ] ],
                        [ 200.0; 150.0 ],
                        style = { TableStyleRef.Default with Name = "MyTableStyle" },
                        cellMargins = { Top = Some 4.0; Bottom = Some 4.0; Left = Some 6.0; Right = Some 6.0 }
                    ) ] ]
        |> withTableStyles [ customStyle ]

    verifyScenarioNamed "Table_CustomStyleAndHeaderRow" "output.docx" doc

// --- Reader resilience against foreign files --------------------------------------------
//
// Not a `verifyScenarioNamed` scenario like the ones above - there's no DSL-level way to
// author a content control to begin with (this DSL doesn't model them, see MAPPING.md),
// so this builds a file directly against the OOXML SDK, bypassing `Writer` entirely, to
// exercise `Reader` against exactly the kind of foreign construct a real-world template
// uses constantly.

[<Fact>]
let ``Reader tolerates unmodeled body-level content instead of throwing`` () =
    let dir = Path.Combine(examplesDir, "ReaderTolerance")
    Directory.CreateDirectory(dir) |> ignore
    let path = Path.Combine(dir, "sdt-and-custom-xml.docx")

    let textRun (text: string) =
        let r = Wordprocessing.Run()
        r.AppendChild(Wordprocessing.Text(text)) |> ignore
        r

    let textParagraph (text: string) =
        let p = Wordprocessing.Paragraph()
        p.AppendChild(textRun text) |> ignore
        p

    do
        use wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document)
        let mainPart = wordDoc.AddMainDocumentPart()
        let body = Wordprocessing.Body()

        body.AppendChild(textParagraph "Before the content control.") |> ignore

        // A content control (`w:sdt`) wrapping a paragraph - this used to make `Document.
        // load` throw outright on the whole file, since `Reader` recognized only `w:p`/
        // `w:tbl` as direct body children.
        let sdtContent = Wordprocessing.SdtContentBlock()
        sdtContent.AppendChild(textParagraph "Inside the content control.") |> ignore
        let sdt = Wordprocessing.SdtBlock()
        sdt.AppendChild(sdtContent) |> ignore
        body.AppendChild(sdt) |> ignore

        // `w:customXml` wrapping a paragraph - same "unmodeled wrapper, recover the real
        // content inside" treatment.
        let customXml = Wordprocessing.CustomXmlBlock()
        customXml.AppendChild(textParagraph "Inside the custom XML wrapper.") |> ignore
        body.AppendChild(customXml) |> ignore

        // `w:altChunk` - an embedded foreign document format this DSL has no way to parse
        // at all, so this one really is dropped rather than recovered; the point here is
        // only that it doesn't take the rest of the file down with it.
        body.AppendChild(Wordprocessing.AltChunk(Id = StringValue "doesNotExist")) |> ignore

        body.AppendChild(textParagraph "After the content control.") |> ignore
        body.AppendChild(Wordprocessing.SectionProperties()) |> ignore

        let documentEl = Wordprocessing.Document()
        documentEl.AppendChild(body) |> ignore
        mainPart.Document <- documentEl

    let doc = Document.load path

    let texts =
        doc.Sections
        |> List.collect (fun s -> s.Body)
        |> List.choose (function
            | ParagraphBlock p -> Some p
            | _ -> None)
        |> List.collect (fun p -> p.Inlines)
        |> List.choose (function
            | Run(text, _, _) -> Some text
            | _ -> None)

    Assert.Equal<string list>(
        [ "Before the content control."; "Inside the content control."; "Inside the custom XML wrapper."; "After the content control." ],
        texts
    )

// --- Slow: actually execute every generated script.fsx and verify it reproduces the
// committed example ------------------------------------------------------------------

[<Trait("Category", "Slow")>]
[<Theory>]
[<InlineData("BasicParagraphsAndRuns")>]
[<InlineData("NamedStyles")>]
[<InlineData("ParagraphFormatting")>]
[<InlineData("BulletList")>]
[<InlineData("NumberedList")>]
[<InlineData("MultiLevelNumberedList")>]
[<InlineData("Table_Basic")>]
[<InlineData("Table_MergedCells")>]
[<InlineData("Table_BordersAndStyle")>]
[<InlineData("Image")>]
[<InlineData("Hyperlink_External")>]
[<InlineData("Bookmark")>]
[<InlineData("Bookmark_MultiParagraph")>]
[<InlineData("Hyperlink_Internal")>]
[<InlineData("Comments")>]
[<InlineData("Comments_MultiParagraph")>]
[<InlineData("TrackedChanges")>]
[<InlineData("FootnotesAndEndnotes")>]
[<InlineData("PageSetupLandscape")>]
[<InlineData("HeaderFooterDefault")>]
[<InlineData("HeaderFooterFirstPageDifferent")>]
[<InlineData("MultipleSections")>]
[<InlineData("DocumentProtectionReadOnly")>]
[<InlineData("Macro")>]
[<InlineData("DocumentProperties")>]
[<InlineData("Table_CustomStyleAndHeaderRow")>]
let ``Regenerated script reproduces the example`` (name: string) =
    let dir = Path.Combine(examplesDir, name)
    let scriptPath = Path.Combine(dir, "script.fsx")
    Assert.True(File.Exists(scriptPath), sprintf "%s doesn't exist yet - run the fast scenario tests first." scriptPath)

    let originalFile = Directory.GetFiles(dir, "output.*") |> Array.filter (fun f -> f.EndsWith(".docx") || f.EndsWith(".docm")) |> Array.exactlyOne

    // Captured before running the script - the script writes to this same filename
    // (that's what "regenerates" means), so the committed file is what gets compared
    // against once the script has overwritten it below.
    let original = Document.load originalFile

    let psi =
        ProcessStartInfo(
            FileName = "dotnet",
            Arguments = sprintf "fsi \"%s\"" scriptPath,
            WorkingDirectory = dir,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        )

    use p = Process.Start(psi)
    let stdout = p.StandardOutput.ReadToEnd()
    let stderr = p.StandardError.ReadToEnd()
    p.WaitForExit()
    Assert.True(p.ExitCode = 0, sprintf "dotnet fsi failed for %s:\n%s\n%s" name stdout stderr)

    let regenerated = Document.load originalFile
    Assert.Equal<Section list>(original.Sections, regenerated.Sections)
