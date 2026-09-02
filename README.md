# Kookerella.FsWordDsl

A typesafe F# DSL for building Word documents, interpreted into calls against the
[DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) SDK. The DSL is a plain
data model (records/DUs with structural equality) - the interpreter (`Writer`) compiles it
to OOXML, and the reverse transform (`Reader`) parses an existing `.docx` back into the
same DSL.

This is the WordprocessingML sibling of
[Kookerella.FsOpenXmlDsl](https://github.com/Kookerella-Ltd/Kookerella.FsOpenXmlDsl) (the
Excel/SpreadsheetML one) - same objectives, same round-trip philosophy, translated to
Word's own document model. See [MAPPING.md](MAPPING.md) for exactly which WordprocessingML
features map 1:1, which are approximated, and which aren't modeled yet.

**This round-trips in both directions**, which most Word libraries don't: they give you an
imperative API to build a document from scratch, but no way to turn an *existing* file back
into readable source. Here, `Reader` parses a real `.docx`/`.docm` back into the same DSL,
and `Document.generateScript` goes one step further and renders that model back out as a
self-contained script that rebuilds an equivalent file - a decompiler for Word documents,
not just a writer. Two more surfaces, `Xml.toDocument`/`Xml.ofDocument` (see ["## XML"](#xml)
below) and `Json.toDocument`/`Json.ofDocument` (see ["## JSON"](#json) below), do the same
translation to/from plain XML or JSON against a real schema - for a caller who'd rather
generate or consume data than write code at all.

**A fluent C# wrapper** (`Kookerella.CsWordDsl`) sits on top of the F# core for callers who'd
rather not touch F# discriminated unions/option types directly - immutable records with
`With*` builders, plus its own `CsCodeGen` decompiler that renders a `Document` back out as
runnable C# source. See ["## The C# wrapper"](#the-c-wrapper) below.

**Scope note**: this repo doesn't ship an MCP server yet (unlike the Excel repo, which has
one alongside its F# core and C# wrapper). See [CLAUDE.md](CLAUDE.md) for what adding one
would look like.

## Layout

- `src/Kookerella.FsWordDsl` - the library.
  - `Units.fs` - conversions between points/inches/pixels and the physical units
    WordprocessingML uses on the wire (twips for page geometry/spacing, EMU for image
    sizing).
  - `Styles.fs` - character and paragraph formatting: `Color` (`Rgb`, `Auto`, or a
    theme-relative `Theme` color - see `ThemeColorKind`), `HighlightColor` (Word's own
    fixed highlight palette), `UnderlineStyle`, `RunStyle` (including small caps/all
    caps/hidden text), `ParagraphAlignment`, `Indentation`, `LineSpacingRule`,
    `TabStopAlignment`/`TabLeader`/`TabStop`, `ParagraphFormat` (including paragraph
    borders, shading, and custom tab stops), `BorderLineStyle`, `BorderSide`,
    `BorderStyle` (reused for both paragraph and table borders).
  - `NamedStyles.fs` - `StyleDefinition` (paragraph or character, with `BasedOn`
    inheritance) and a small `BuiltInStyles` catalog (`normal`, `heading1`/`2`/`3`,
    `title`, `listParagraph`, `hyperlinkCharStyle`).
  - `Numbering.fs` - `NumberFormatKind`, `ListLevel`, `NumberingDefinition` for
    numbered/bulleted lists, including multi-level ones (`ListLevel` isn't limited to one
    per definition - see `Builders.multiLevelNumberedListDef`).
  - `Hyperlinks.fs` - `HyperlinkTarget` (external URL vs. internal bookmark reference).
  - `Protection.fs` - `EditRestriction` and `DocumentProtection`, document-level (Word has
    no per-section equivalent of Excel's per-sheet protection).
  - `Revisions.fs` - `RevisionKind`/`Revision` for track changes (`Inline.TrackedChange`,
    `Paragraph.MarkRevision`) - narrowly scoped to inserted/deleted content and paragraph
    marks, see `MAPPING.md` for what isn't covered.
  - `ContentControls.fs` - `ContentControlType`/`ContentControlProps` for content controls
    (structured document tags, `w:sdt`): plain text, rich text, dropdown/combo box, date
    picker, checkbox - see `MAPPING.md` for what isn't covered.
  - `PageSetup.fs` - `PageOrientation`, `PageSize`, `PageMargins`, `SectionBreakType`,
    `NoteNumberRestart`/`NoteNumberingSettings` (a section's own footnote/endnote
    numbering).
  - `Tables.fs` - `TableBorders`, `VerticalMergeKind`, `TableCellProps` (including a
    per-cell `Margins` override), `TableStyleRef`, `TableStyleRegion`/
    `TableStyleDefinition` (custom table style definitions - all thirteen of OOXML's
    conditional-formatting regions), and `CellMargins` (shared shape for a table's default
    margins and a single cell's own override).
  - `Images.fs` - `ImageFormat`, `ImageEntry` (raw file bytes plus an on-page size),
    anchored inline within a run.
  - `DocumentProperties.fs` - `DocumentProperties` (Title, Author, Subject, Keywords,
    Comments, Category, Company) - core document metadata, `Document.Properties`.
  - `Model.fs` - the recursive content model: `Inline` (runs, breaks, images, hyperlinks,
    bookmarks and comments - both the single-paragraph `Bookmark`/`Comment` cases and the
    cross-paragraph `BookmarkRangeStart`/`End`/`CommentRangeStart`/`End` markers, simple
    fields, footnotes/endnotes, `TrackedChange` for track changes, and
    `InlineContentControl` for content controls), `Paragraph` (including `MarkRevision`),
    `Block` (paragraph, table, or `ContentControlBlock` - the block-level counterpart to
    `InlineContentControl`), `TableCell`/
    `TableRow` (including `RepeatAsHeader`)/`TableEntry` (including `CellMargins`),
    `HeaderFooterSet`,
    `SectionProperties` (including `BreakType` and `FootnoteNumbering`/
    `EndnoteNumbering`), `Section`, `Document` (including `Document.VbaProject`, a
    macro-enabled document's raw `vbaProject.bin` bytes, `Document.Properties`, and
    `Document.TableStyles`).
  - `Xml.fs` / `Xml.xsd` - the XML surface: `Xml.toDocument`/`Xml.ofDocument` translate a
    `Document` to/from an `XElement` tree, and `Xml.schemaSet()` loads the paired schema
    (embedded in the assembly as a resource) for validating either direction. See
    ["## XML"](#xml) below.
  - `Json.fs` / `Json.schema.json` - the JSON surface: `Json.toDocument`/`Json.ofDocument`
    translate a `Document` to/from a `System.Text.Json.Nodes.JsonObject` tree. Schema
    validation is test-suite only, not a public API. See ["## JSON"](#json) below.
  - `Builders.fs` - plain functional constructors (`section`, `document`, `withStyles`,
    `withNumbering`, `withProtection`, `withVbaProject`, `withDocumentProperties`,
    `withTableStyles`, `bulletListDef`, `numberedListDef`, `multiLevelNumberedListDef`)
    plus `DocumentDsl` - smart constructors (`run`, `para` (with `markRevision`),
    `hyperlink`, `bookmark`, `comment`, `inserted`/`deleted` (track changes),
    `contentControl`/`contentControlBlock` (content controls), `image`, `footnote`,
    `endnote`, `tableCell`, `tableRow` (with `height`/`repeatAsHeader`), `table` (with
    `style`/`borders`/`cellMargins`)) with real optional parameters, the Word analog of
    the Excel repo's `SheetDsl`.
  - `Interpreter/StyleRegistry.fs` - shared run/paragraph/border/color conversions plus
    `Document.Styles` <-> `styles.xml` (internal).
  - `Interpreter/ImageWriter.fs` / `ImageReader.fs` - an inline image's own DSL <->
    DrawingML translation (internal).
  - `Interpreter/Writer.fs` - DSL -> OOXML (internal).
  - `Interpreter/Reader.fs` - OOXML -> DSL, the reverse transform (internal).
  - `Interpreter/CodeGen.fs` - DSL -> F# *source text*: renders a `Document` back out as a
    self-contained `.fsx` script that rebuilds an equivalent file when run (internal).
  - `Api.fs` - the public `Document.save`/`saveToStream`/`load`/`loadFromStream`/
    `generateScript` entry points.
- `src/Kookerella.CsWordDsl` - the fluent C# wrapper (see ["## The C# wrapper"](#the-c-wrapper)
  below): immutable records/`sealed record` closed hierarchies mirroring the F# core's own
  types one-for-one, `DocumentConverter.cs` (internal, the two-way F#<->C# translation),
  `DocumentIO.cs` (`Save`/`Load`, the one place this project does I/O), `CsCodeGen.cs` (DSL
  -> C# *source text*, the C# analog of `Interpreter/CodeGen.fs`).
- `tests/Kookerella.FsWordDsl.Tests` - one test per feature, each validating the produced
  file against the OOXML schema (`DocumentFormat.OpenXml.Validation.OpenXmlValidator`) and
  asserting an exact round trip back through the DSL. Each test also writes the document it
  builds to `Examples/<test name>/output.docx` (checked into the repo), plus `script.fsx`
  (regenerates the file - a separate, slower `Category=Slow` test group actually executes
  each one via `dotnet fsi`), `document.xml`, and `document.json` - one folder always has
  four views of the same example.
- `tests/Kookerella.CsWordDsl.Tests` - `DriftGuardTests.cs` (a reflection-based tripwire
  checking the C# wrapper's DU mirrors haven't fallen behind the F# core's own case
  counts), `DocumentTests.cs` (targeted round-trip assertions per feature),
  `ExampleTests.cs` (reloads the F# suite's own checked-in `Examples/*/output.docx`
  fixtures rather than re-authoring every scenario a second time), `CsCodeGenTests.cs`
  (actually executes a generated file via `dotnet run --file`, the C# analog of the F#
  suite's `Category=Slow` `dotnet fsi` group).
- `samples/Kookerella.FsWordDsl.Sample` - a small console app that builds a document, saves
  it, and reads it back.

## Quick start

```fsharp
open Kookerella.FsWordDsl
open type Kookerella.FsWordDsl.DocumentDsl

let doc =
    document
        [ section
              [ para ([ run "Quarterly Report" ], styleId = "Title")
                para
                    [ run "This report covers "
                      run ("Q1 2026", style = { RunStyle.Default with Bold = true })
                      run ", see the "
                      hyperlink ("full dataset", ExternalUrl "https://example.com/data")
                      run " for details." ] ] ]

doc |> Document.save "report.docx"

// Reverse transform:
let roundTripped = Document.load "report.docx"
```

`document` defaults `Styles` to `BuiltInStyles.all`, so `styleId = "Heading1"` (or any other
built-in id) just works without registering it first - pipe `withStyles` afterward to
replace or extend that set. `run`/`para`/`hyperlink`/`bookmark`/`comment`/`image`/
`tableCell`/`tableRow`/`table` are `DocumentDsl` members with real optional parameters
(`open type Kookerella.FsWordDsl.DocumentDsl` brings them into scope unqualified, same as
`open type SheetDsl` does in the Excel repo) - plain F# `let` bindings can't have optional
parameters, which is why this part of the DSL is a type.

A `Paragraph`'s `Inlines` are naturally several independently-styled runs - rich text
(mixed formatting within one paragraph) is first-class, not a documented gap the way
Excel's single-uniform-run `Text` cell is:

```fsharp
para
    [ run "Plain text, "
      run ("bold", style = { RunStyle.Default with Bold = true })
      run ", and "
      run ("colored", style = { RunStyle.Default with Color = Some Color.red }) ]
```

`RunStyle` also covers small caps, all caps, and hidden text; `ParagraphFormat` covers
borders (`BorderStyle`, the same shape used for table borders) and shading:

```fsharp
para
    ([ run "ALL CAPS AND SMALL CAPS" ], format =
        { ParagraphFormat.Default with
            Borders = Some { BorderStyle.None with Bottom = Some { Style = SingleLine; Width = Some 1.0; Color = Some Color.black } }
            Shading = Some(Rgb(0xD9uy, 0xD9uy, 0xD9uy)) })
```

Custom tab stops (`TabStop`) sit on `ParagraphFormat.TabStops` - a right-aligned stop with a
dot leader is the classic table-of-contents pattern:

```fsharp
para
    ([ run "Introduction"; Tab; run "1" ], format =
        { ParagraphFormat.Default with TabStops = [ { Position = 288.0; Alignment = RightTab; Leader = DotLeader } ] })
```

`Color` also accepts a theme-relative token (`Theme`) alongside plain `Rgb`/`Auto` - since
this DSL has no theme part to resolve it against, real Word does that; `Fallback` is what a
themeless reader sees instead, the same "always also write a computed value" convention Word
itself follows:

```fsharp
run ("Accent-colored text", style = { RunStyle.Default with Color = Some(Theme(Accent1Theme, (0x1Fuy, 0x49uy, 0x7Duy), None, None)) })
```

Lists use a `(numId, level)` reference on the paragraph, resolved against a
`NumberingDefinition` attached to the document - `NumberingDefinition.Levels` isn't limited
to one level, and `multiLevelNumberedListDef` builds the common correctly-linked outline
shape for you:

```fsharp
document
    [ section
          [ para ([ run "First bullet" ], numbering = (1, 0))
            para ([ run "Second bullet" ], numbering = (1, 0)) ] ]
|> withNumbering [ bulletListDef 1 ]

document
    [ section
          [ para ([ run "First topic" ], numbering = (1, 0))
            para ([ run "First subtopic" ], numbering = (1, 1))
            para ([ run "Second topic" ], numbering = (1, 0)) ] ]
|> withNumbering [ multiLevelNumberedListDef 1 3 ]
```

Tables are built from `tableRow`/`tableCell`, with column widths given once for the whole
table - a cell without an explicit width falls back to its column's width at write time:

```fsharp
table (
    [ tableRow [ tableCell [ para [ run "Item" ] ]; tableCell [ para [ run "Qty" ] ] ]
      tableRow [ tableCell [ para [ run "Widgets" ] ]; tableCell [ para [ run "12" ] ] ] ],
    [ 200.0; 100.0 ],
    style = TableStyleRef.Default
)
```

Cell merging - horizontal (`GridSpan`) and vertical (`RestartMerge`/`ContinueMerge`) - are
independent and combine on the same cell, matching real Word:

```fsharp
tableCell ([ para [ run "Spans 2 columns" ] ], props = { TableCellProps.Default with GridSpan = Some 2 })
```

A cell's own margins override the table's default the same `CellMargins` shape covers both:

```fsharp
tableCell ([ para [ run "Extra padding" ] ], props = { TableCellProps.Default with Margins = Some { CellMargins.Default with Top = Some 8.0; Bottom = Some 8.0 } })
```

A custom table style (`TableStyleDefinition`) lives in `Document.TableStyles` and is applied
by name, the same way a built-in like `"TableGrid"` is - here with a bold white header row on
a blue background, an italic last row, and alternating row shading, plus a table-wide default
cell margin and a row that repeats on every page:

```fsharp
let corporateStyle: TableStyleDefinition =
    { TableStyleDefinition.Default with
        Id = "Corporate"
        Name = "Corporate"
        FirstRow =
            { TableStyleRegion.None with
                RunFormat = Some { RunStyle.Default with Bold = true; Color = Some Color.white }
                CellShading = Some(Rgb(0x4Fuy, 0x81uy, 0xBDuy)) }
        LastRow = { TableStyleRegion.None with RunFormat = Some { RunStyle.Default with Italic = true } }
        BandedRow = { TableStyleRegion.None with CellShading = Some(Rgb(0xDCuy, 0xE6uy, 0xF1uy)) } }

document
    [ section
          [ table (
                [ tableRow ([ tableCell [ para [ run "Item" ] ]; tableCell [ para [ run "Qty" ] ] ], repeatAsHeader = true)
                  tableRow [ tableCell [ para [ run "Widgets" ] ]; tableCell [ para [ run "12" ] ] ] ],
                [ 200.0; 100.0 ],
                style = { TableStyleRef.Default with Name = "Corporate" },
                cellMargins = { Top = Some 4.0; Bottom = Some 4.0; Left = Some 6.0; Right = Some 6.0 }
            ) ] ]
|> withTableStyles [ corporateStyle ]
```

`TableStyleDefinition` also covers `FirstColumn`/`LastColumn`, `BandedColumn`, and the four
corner cells (`NorthEastCell`/`NorthWestCell`/`SouthEastCell`/`SouthWestCell`) - the two
regions not modeled are each banding axis's *second* band, since in practice that's just
`WholeTable`'s own background showing through (see [MAPPING.md](MAPPING.md)).

Sections carry their own page setup - a document is a sequence of `Section`s, mapping 1:1
onto real Word section breaks. `BreakType` is how a section begins *relative to the
previous one* - meaningless (and not written) on the very first section:

```fsharp
let landscape = { SectionProperties.Default with Orientation = Landscape }
document [ sectionWith landscape [ para [ run "A landscape-oriented page." ] ] ]

let continuous = { SectionProperties.Default with BreakType = ContinuousBreak }
document
    [ section [ para [ run "Section 1." ] ]
      sectionWith continuous [ para [ run "Section 2 - no page break from section 1." ] ] ]
```

Footnotes and endnotes mark a point in a paragraph's own `Inlines` - `content` is the
note's own body, written to `word/footnotes.xml`/`endnotes.xml` with an id `Writer` assigns
automatically (the reference-mark run itself is generated for you, on both ends):

```fsharp
para
    [ run "This claim needs a citation"
      footnote "Smith, J. (2023). A Study of Claims."
      run ", and this one refers to a fuller discussion"
      endnote [ para [ run "See the appendix for the full derivation." ] ] ]
```

A section's own footnote/endnote numbering (`w:footnotePr`/`w:endnotePr`) - `None` is Word's
own default (continuous decimal from 1); here footnotes are lower-roman and restart every
page, matching a common legal-document convention:

```fsharp
sectionWith
    { SectionProperties.Default with FootnoteNumbering = Some { Format = LowerRomanFormat; StartAt = None; Restart = RestartEachPage } }
    [ para [ run "Body text."; footnote "A footnote numbered i, ii, iii, ... restarting each page." ] ]
```

Headers and footers are per-section, with `Default`/`First`/`Even` variants (the
`titlePg`/`evenAndOddHeaders` flags real Word needs are set automatically):

```fsharp
let footer = { HeaderFooterSet.None with Default = Some [ para [ run "Page "; Field("PAGE", Some "1") ] ] }
sectionWith { SectionProperties.Default with Footer = Some footer } [ para [ run "Body text." ] ]
```

Comments and bookmarks wrap inline content directly, the common single-paragraph case:

```fsharp
para [ comment ([ run "This figure needs review." ], "Please double check the totals.", author = "Alex") ]
```

Either spanning more than one paragraph uses two independent markers placed directly in
separate paragraphs instead, sharing an id - `BookmarkRangeStart`/`BookmarkRangeEnd` for
bookmarks, `CommentRangeStart`/`CommentRangeEnd` for comments (which carries the comment's
own metadata on its `Start`, since there's no wrapping case here to hang it off - see
[MAPPING.md](MAPPING.md) on why that id is write-time-only, unlike a bookmark's own name):

```fsharp
document
    [ section
          [ para [ BookmarkRangeStart "Section2"; run "This paragraph starts the bookmark" ]
            para [ run "and this one ends it."; BookmarkRangeEnd "Section2" ] ] ]

document
    [ section
          [ para [ CommentRangeStart("review1", "Alex", None, None, "This section needs review."); run "Comment starts here" ]
            para [ run "and ends here."; CommentRangeEnd "review1" ] ] ]
```

Track changes (`inserted`/`deleted`) wrap inline content the same way, marking it as
inserted or deleted under an author and date; a whole inserted or deleted paragraph
(rather than just some of its content) uses `para`'s own `markRevision` instead, for the
paragraph's closing mark:

```fsharp
para
    [ run "The quick "
      inserted ([ run "brown " ], "Alex")
      run "fox jumps over the "
      deleted ([ run "lazy " ], "Alex")
      run "dog." ]

para ([ run "This whole paragraph was inserted." ], markRevision = { Kind = Inserted; Author = "Alex"; Date = None })
```

Content controls (`contentControl` for run-level, `contentControlBlock` for block-level)
wrap their own currently-displayed content the same way, plus a `ContentControlType` (see
[MAPPING.md](MAPPING.md) for the full set - plain text, rich text, dropdown/combo box, date
picker, checkbox):

```fsharp
para
    [ run "Client name: "
      contentControl ([ run "Type here" ], PlainTextControl false, alias = "Client Name", tag = "clientName") ]

para
    [ run "Favorite color: "
      contentControl ([ run "Blue" ], DropDownControl([ "Red", "red"; "Green", "green"; "Blue", "blue" ], false)) ]

contentControlBlock([ para [ run "This whole paragraph is a rich-text content control." ] ], RichTextControl, alias = "Notes")
```

Document-level protection, macros, and core properties are all pipe-friendly, same shape as
Excel's own `withProtection`/`withVbaProject`:

```fsharp
document [...] |> withProtection { Edit = Some ReadOnlyRestriction; Password = Some "hunter2" }
document [...] |> withVbaProject (System.IO.File.ReadAllBytes("vbaProject.bin"))
document [...] |> withDocumentProperties { DocumentProperties.Default with Title = Some "Quarterly Report"; Author = Some "Kookerella" }
```

Save the result with a `.docm` path - `Document.save`/`saveToStream` automatically switch
the file's own declared content type to Word's macro-enabled kind whenever a `VbaProject`
is present, but real Word also expects the `.docm` extension to trust and run macros at all.

## Regenerating a file as F# source

Given a `Document` (typically one you just `Document.load`ed from an existing file),
`Document.generateScript` renders it back out as a self-contained `.fsx` script that
rebuilds an equivalent file when run - a code-generating counterpart to `Document.load`:

```fsharp
let doc = Document.load "input.docx"

let referenceLines =
    [ "#r \"path/to/Kookerella.FsWordDsl.dll\""
      "#r \"path/to/DocumentFormat.OpenXml.dll\"" ]

let script = Document.generateScript referenceLines "output.docx" doc
System.IO.File.WriteAllText("regenerate.fsx", script)
```

Running `dotnet fsi regenerate.fsx` produces `output.docx` - not byte-identical to the
original (zip metadata/timestamps differ) but structurally equivalent through the same
round-trip lens every other test in this repo uses. Every scenario under `tests/
Kookerella.FsWordDsl.Tests/Examples/` has a committed `script.fsx` generated exactly this
way; the `Category=Slow` test group actually executes each one via `dotnet fsi` and checks
it reproduces the committed `.docx`.

## The C# wrapper

`Kookerella.CsWordDsl` is an idiomatic, immutable, fluent C# wrapper over the F# core, for
callers who'd rather not touch F# discriminated unions or option types directly. Every F#
type has a C# mirror: plain records with `With*`/factory-method builders for product types,
`enum`s for parameterless choices, and `sealed record` closed hierarchies (`abstract record`
base, private constructor, nested cases) for everything else - the same "sealed hierarchy"
pattern the Excel repo's own `Kookerella.CsOpenXmlDsl` uses for `CellValue`/
`ConditionalFormatRule`. Reference `Kookerella.CsWordDsl` instead of `Kookerella.FsWordDsl`
and never see an `FSharpOption`:

```csharp
using Kookerella.CsWordDsl;

var doc = Document.Create(
    Section.Of([
        Block.Paragraph([new Inline.Run("Quarterly Report")], styleId: "Title"),
        Block.Paragraph([
            new Inline.Run("This report covers "),
            new Inline.Run("Q1 2026", new RunStyle { Bold = true }),
            Inline.HyperlinkText("full dataset", new HyperlinkTarget.ExternalUrl("https://example.com/data")),
            new Inline.Run(" for details.")
        ])
    ]));

DocumentIO.Save(doc, "report.docx");
var loaded = DocumentIO.Load("report.docx");
```

Content controls, tables, track changes, comments, and every other feature the F# core
models are covered the same way - see `tests/Kookerella.CsWordDsl.Tests/DocumentTests.cs`
for a worked example per feature. `CsCodeGen.Generate` is the C# analog of
`Document.generateScript`: it renders a `Document` back out as a self-contained C# file
targeting .NET's "file-based apps" feature (`dotnet run --file script.cs`), rather than an
`.fsx` script:

```csharp
var script = CsCodeGen.Generate(["#:project path/to/Kookerella.CsWordDsl.csproj"], "output.docx", loaded);
File.WriteAllText("regenerate.cs", script);
```

One design note worth stating explicitly: this wrapper's records use `IReadOnlyList<T>`
properties, and C#'s compiler-synthesized record equality does not deep-compare list
contents (two records holding equal-but-distinct list instances compare unequal via plain
`.Equals()`) - the same limitation `Kookerella.CsOpenXmlDsl`'s own records have. Don't rely
on whole-`Document` equality in your own code; compare the specific values you care about,
the same way this repo's own `DocumentTests.cs` does.

## XML

`Xml.toDocument`/`Xml.ofDocument` (in `Xml.fs`) are a third way in and out of the DSL,
alongside writing F# directly and code generation: plain XML, against a real schema
(`Xml.xsd`, embedded in the assembly). A data-carrying DU case becomes an element named
after the case; a parameterless-choice case becomes an attribute value or bare string,
matching the convention the Excel repo's own `Xml.fs` documents.

```fsharp
open System.Xml.Linq

// XML -> Document -> .docx
let doc = XElement.Load "report.xml" |> Xml.ofDocument
Document.save "report.docx" doc

// .docx -> Document -> XML
let xml = Document.load "report.docx" |> Xml.toDocument
xml.Save "report.xml"
```

A run with direct formatting and a hyperlink, in XML:

```xml
<para>
  <run>Visit </run>
  <hyperlink tooltip="Kookerella on GitHub">
    <externalHyperlink>https://github.com/Kookerella-Ltd</externalHyperlink>
    <content>
      <run styleId="Hyperlink">Kookerella on GitHub</run>
    </content>
  </hyperlink>
  <run> for more.</run>
</para>
```

`Xml.schemaSet()` loads the compiled schema for validating either direction yourself
(`XDocument.Validate`) - every scenario under `tests/Kookerella.FsWordDsl.Tests/Examples/`
has a committed `document.xml` validated against it this way as part of the same test that
generates it.

## JSON

`Json.toDocument`/`Json.ofDocument` (in `Json.fs`) are a fourth way in and out of the DSL,
alongside writing F# directly, code generation, and XML: plain JSON, for a caller whose
tooling speaks JSON rather than XML. The same DU-case conventions apply, in JSON's own
idiom (a single-key object for a data-carrying case, a bare string for a parameterless one):

```fsharp
open System.Text.Json.Nodes

// JSON -> Document -> .docx
let doc = JsonNode.Parse(File.ReadAllText "report.json").AsObject() |> Json.ofDocument
Document.save "report.docx" doc

// .docx -> Document -> JSON
let json = Document.load "report.docx" |> Json.toDocument
File.WriteAllText("report.json", json.ToJsonString())
```

The same hyperlink example as above, in JSON:

```json
{
  "para": {
    "inlines": [
      { "run": { "text": "Visit " } },
      {
        "hyperlink": {
          "target": { "externalHyperlink": "https://github.com/Kookerella-Ltd" },
          "runs": [ { "run": { "text": "Kookerella on GitHub", "styleId": "Hyperlink" } } ],
          "tooltip": "Kookerella on GitHub"
        }
      },
      { "run": { "text": " for more." } }
    ]
  }
}
```

Unlike XML, .NET has no built-in JSON Schema validator, so `Json.schema.json` (in the repo)
is validated only from this repo's own test suite (via a test-only `JsonSchema.Net`
dependency) rather than exposed as a public API - see `Json.fs`'s own doc comment.

## Building and testing

```bash
dotnet build
dotnet test --filter "Category!=Slow"
dotnet run --project samples/Kookerella.FsWordDsl.Sample
```

The default loop above skips the slow `Category=Slow` tests, which actually invoke
`dotnet fsi` on every generated `Examples/*/script.fsx` (multi-second process startup each).
Run those explicitly, after the fast suite has populated the `.fsx` files at least once:

```bash
dotnet test --filter "Category=Slow"
```

Plain `dotnet test` (no filter) runs both groups.

The C# wrapper's own suite has no fast/slow split - `CsCodeGenTests.cs` shells out to
`dotnet run --file` itself, so a single run already covers the C# analog of the F# suite's
slow group:

```bash
dotnet test tests/Kookerella.CsWordDsl.Tests
```
