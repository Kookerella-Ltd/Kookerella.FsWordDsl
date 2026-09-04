# Kookerella.CsWordDsl

An idiomatic, immutable, fluent C# wrapper over
[Kookerella.FsWordDsl](https://www.nuget.org/packages/Kookerella.FsWordDsl) - build and read
Word documents (.docx/.dotm) from C# without touching F# discriminated unions or option
types directly.

Every F# type has a C# mirror: plain `record`s with `With*`/factory-method builders for
product types, `enum`s for parameterless choices, and `sealed record` closed hierarchies
(`abstract record` base, private constructor, nested cases) for everything else - the same
"sealed hierarchy" pattern the Excel repo's own `Kookerella.CsOpenXmlDsl` uses for
`CellValue`/`ConditionalFormatRule`. The only place this library does any I/O at all is
`DocumentIO` - every other type is pure data.

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

A `Paragraph`'s `Inlines` naturally hold several independently-styled `Inline.Run`s, so rich
text (mixed formatting within one paragraph) is first-class here - unlike a spreadsheet
cell's single uniform value. `Document.Create` starts a document with `BuiltInStyles.All`
already registered, so `StyleId = "Heading1"` (or any other built-in id) just works without
registering it first.

Bulleted and numbered lists are a `NumberingDefinition` on the `Document`, referenced from a
paragraph by `(numId, level)`:

```csharp
var doc = Document.Create(
    Section.Of([
        Block.Paragraph([new Inline.Run("Widgets")], numbering: (1, 0)),
        Block.Paragraph([new Inline.Run("Gadgets")], numbering: (1, 0))
    ])).WithNumbering(NumberingDefinition.BulletList(1));
```

`NumberingDefinition.NumberedList`/`.MultiLevelNumberedList` cover the other common shapes; a
custom list (a different glyph, a Roman-numeral level, a bespoke indent) is built directly
from `ListLevel` entries instead.

Tables are rows of cells, addressed sequentially rather than by a sparse row/column cursor -
`GridSpan`/`VerticalMerge` are the only place a cell's position needs to be stated explicitly:

```csharp
var table = Block.Table(
    [
        TableRow.Of([TableCell.Of([Block.Paragraph([new Inline.Run("Item")])]),
                     TableCell.Of([Block.Paragraph([new Inline.Run("Qty")])])]),
        TableRow.Of([TableCell.Of([Block.Paragraph([new Inline.Run("Widgets")])]),
                     TableCell.Of([Block.Paragraph([new Inline.Run("42")])])])
    ],
    columnWidths: [225.0, 75.0],
    style: TableStyleRef.Named("TableGrid") with { FirstRowBanding = true });
```

`TableCellProps.WithGridSpan`/`.WithVerticalMerge` combine independently on the same cell for
merged regions; `TableBorders`/`CellMargins` set explicit gridlines and spacing when a named
style isn't enough.

Images are embedded inline within a run, sized in EMU or, more conveniently, inches - this
wrapper does no decoding of its own, `Data` is exactly the bytes of the image file on disk,
handed back unchanged on read:

```csharp
var image = new Inline.Image(
    ImageEntry.FromBytesInches(File.ReadAllBytes("logo.png"), ImageFormat.Png, 2.0, 1.0, altText: "Company logo"));
```

Bookmarks, comments, footnotes, and endnotes each come in two shapes: a single-paragraph
convenience case, and a `*RangeStart`/`*RangeEnd` pair for content spanning more than one
paragraph:

```csharp
var doc = Document.Create(
    Section.Of([
        Block.Paragraph([
            new Inline.Bookmark("intro", [new Inline.Run("Introduction")]),
            new Inline.Run(" - see also "),
            new Inline.Hyperlink(new HyperlinkTarget.InternalBookmark("intro"), [new Inline.Run("here")])
        ]),
        Block.Paragraph([
            new Inline.Run("This claim needs a citation."),
            new Inline.Comment("Alex", "AR", null, "Please add a source.", [new Inline.Run("needs a citation")])
        ]),
        Block.Paragraph([
            new Inline.Run("A footnoted fact."),
            Inline.FootnoteText("Source: internal survey, 2026.")
        ])
    ]));
```

Track changes wrap arbitrary inline content, marking it inserted or deleted by whoever made
the change and when:

```csharp
var para = Block.Paragraph([
    new Inline.Run("The deadline is "),
    Inline.Deleted([new Inline.Run("Friday")], author: "Alex"),
    Inline.Inserted([new Inline.Run("Monday")], author: "Alex"),
    new Inline.Run(".")
]);
```

Content controls (structured document tags) come in five kinds - plain text, rich text,
drop-down/combo box, date, and checkbox - and can wrap either inline content within a
paragraph (`Inline.ContentControl`) or whole paragraphs/tables (`Block.ContentControlBlock`):

```csharp
var checkbox = new Inline.ContentControl(
    new ContentControlProps { Type = new ContentControlType.CheckBox(Checked: false) }.WithTag("agree"),
    [new Inline.Run(" ")]);

var dropdown = new Inline.ContentControl(
    new ContentControlProps
    {
        Type = new ContentControlType.DropDown([("Low", "low"), ("Medium", "medium"), ("High", "high")])
    }.WithAlias("Priority"),
    [new Inline.Run("Medium")]);
```

`ContentControlProps.WithLock` restricts editing (`LockDeletion`, `LockContentEditing`, or
both) the same way Word's own "Content control cannot be deleted/edited" checkboxes do.

Print/page setup - size, orientation, margins, headers/footers, columns - lives on each
`Section`'s own `SectionProperties`, since a document is a sequence of sections rather than
one flat page setup, matching real Word section breaks:

```csharp
var section = Section.With(
    SectionProperties.Default
        .WithOrientation(PageOrientation.Landscape)
        .WithPageSize(new PageSize.A4())
        .WithMargins(PageMargins.Default with { Left = 36.0, Right = 36.0 })
        .WithFooter(HeaderFooterSet.None.WithDefault(
            [Block.Paragraph([new Inline.Field("PAGE", "1")])])),
    body: [Block.Paragraph([new Inline.Run("Wide table below")])]);
```

`HeaderFooterSet`'s `Default`/`First`/`Even` mirror Word's own three header/footer variants
exactly - `First` shows only on a section's first page, `Even` only on even pages.

Document-level protection is a single mutually-exclusive edit restriction, `null` by default
(unprotected) - Word has no per-section equivalent of a spreadsheet's per-sheet protection:

```csharp
var protectedDoc = doc.WithProtection(DocumentProtection.With(EditRestriction.CommentsOnly));
```

A VBA project (macros) is opaque bytes, same treatment as the F# core and as the Excel
sibling's own `WorkbookProtection` - nothing in this stack parses, generates, or edits VBA
source, it only embeds and hands back exactly what you give it:

```csharp
var macroDoc = doc.WithVbaProject(File.ReadAllBytes("vbaProject.bin"));
DocumentIO.Save(macroDoc, "out.dotm"); // .dotm, not .docx - see below
```

Save to a `.docm`/`.dotm` path once a VBA project is attached - the file's content type
switches to macro-enabled automatically, but real Word also expects the extension to match
before it will trust and run macros regardless of what the content type says.

`CsCodeGen.Generate` is the C# analog of `Document.generateScript`: it renders a `Document`
back out as a self-contained C# file targeting .NET's "file-based apps" feature (`dotnet run
--file script.cs`), rather than an `.fsx` script - the reverse of `DocumentIO.Load` one level
further: loading turns a file into these types, this turns those types into C# *source text*:

```csharp
var script = CsCodeGen.Generate(["#:project path/to/Kookerella.CsWordDsl.csproj"], "output.docx", loaded);
File.WriteAllText("regenerate.cs", script);
// then: dotnet run --file regenerate.cs
```

The first argument is whatever raw `#:package`/`#:project` directive lines the emitted file
needs to locate this assembly (pass a `#:package Kookerella.CsWordDsl@0.1.0` line instead
when generating against the published package rather than a local checkout). Generated code
only mentions what isn't already implied by a type's own defaults, so it reads close to how a
human would write it by hand.

`DocumentIO` also reaches the F# core's other two ways in and out - XML and JSON, each
against a real schema - and F# script generation itself, so a C# caller never needs its own
`Kookerella.FsWordDsl` reference to get any of them:

```csharp
var xml = DocumentIO.ToXml(loaded);   // matches the embedded Xml.xsd
var fromXml = DocumentIO.FromXml(xml);

var json = DocumentIO.ToJson(loaded); // matches Json.schema.json's shape
var fromJson = DocumentIO.FromJson(json);

var fsScript = DocumentIO.GenerateFSharpScript(["#r \"path/to/Kookerella.FsWordDsl.dll\""], "output.docx", loaded);
```

`GenerateFSharpScript` is `CsCodeGen.Generate`'s F#-targeting sibling - same "hand back
runnable source, not just data" idea, just producing an `.fsx` for `dotnet fsi` instead of a
C# file. `ToXml`/`FromXml`/`ToJson`/`FromJson` are thin wrappers over the F# core's own
`Xml.fs`/`Json.fs` - no translation happens in this assembly beyond the same F#↔C# shape
conversion `Save`/`Load` already do, so the XML/JSON is byte-for-byte what
`Kookerella.FsWordDsl` itself would produce.

See [`tests/Kookerella.CsWordDsl.Tests/DocumentTests.cs`](../../tests/Kookerella.CsWordDsl.Tests/DocumentTests.cs)
for a worked round-trip example per feature above, and
[`tests/Kookerella.CsWordDsl.Tests/ExampleTests.cs`](../../tests/Kookerella.CsWordDsl.Tests/ExampleTests.cs)
for every scenario the F# core's own `Examples/` gallery covers, reloaded through this
wrapper.

## Scope

This wrapper covers every feature the F# core models: paragraphs and runs (including mixed
formatting within one paragraph), named styles, numbered/bulleted lists, tables (merges,
borders, custom table style definitions), images, hyperlinks, bookmarks, comments, sections
and page setup, headers/footers (default/first/even), document protection, footnotes/
endnotes, track changes, content controls (all five kinds), VBA (as opaque bytes), and all
four ways in and out the F# core itself offers: `Save`/`Load`, C# code generation
(`CsCodeGen`), F# code generation (`DocumentIO.GenerateFSharpScript`), and schema-backed
XML/JSON (`DocumentIO.ToXml`/`FromXml`/`ToJson`/`FromJson`).

One design note worth stating explicitly: this wrapper's records use `IReadOnlyList<T>`
properties, and C#'s compiler-synthesized record equality does not deep-compare list
contents (two records holding equal-but-distinct list instances compare unequal via plain
`.Equals()`) - the same limitation `Kookerella.CsOpenXmlDsl`'s own records have. Don't rely
on whole-`Document` equality in your own code; compare the specific values you care about,
the same way this repo's own `DocumentTests.cs` does.
