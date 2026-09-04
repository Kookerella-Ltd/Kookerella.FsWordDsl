module XmlTests

open System.Xml.Linq
open System.Xml.Schema
open Xunit
open Kookerella.FsWordDsl

/// Validates `doc` against the embedded `Xml.xsd` - shared by `Tests.verifyScenarioNamed`
/// (every scenario's `document.xml`) and this file's own direct round-trip tests.
let assertXmlSchemaValid (doc: XDocument) =
    let schemaSet = Xml.schemaSet ()
    let mutable errors = []
    doc.Validate(schemaSet, (fun _ e -> errors <- e.Message :: errors))
    Assert.True(errors.IsEmpty, String.concat "\n" errors)

let private minimalDocument: Document =
    document
        [ section
              [ ParagraphBlock
                    { Inlines = [ Run("Hello, World!", Some { RunStyle.Default with Bold = true }, None) ]
                      StyleId = None
                      Format = None
                      Numbering = None
                      MarkRevision = None } ] ]

[<Fact>]
let ``Xml round trips a minimal document`` () =
    let xml = Xml.toDocument minimalDocument
    let roundTripped = Xml.ofDocument xml
    Assert.Equal<Section list>(minimalDocument.Sections, roundTripped.Sections)
    assertXmlSchemaValid (XDocument(xml))

[<Fact>]
let ``Xml round trips styles, numbering, and protection`` () =
    let doc =
        minimalDocument
        |> withStyles [ BuiltInStyles.normal; BuiltInStyles.heading1 ]
        |> withNumbering [ bulletListDef 1 ]
        |> withProtection { Edit = Some ReadOnlyRestriction; Password = None }

    let xml = Xml.toDocument doc
    let roundTripped = Xml.ofDocument xml
    // Sorted before comparing on both sides - toDocument itself sorts Styles/Numbering by
    // Id (see its own doc comment), so content, not input order, is what this asserts.
    Assert.Equal<StyleDefinition list>(doc.Styles |> List.sortBy (fun s -> s.Id), roundTripped.Styles |> List.sortBy (fun s -> s.Id))
    Assert.Equal<NumberingDefinition list>(doc.Numbering |> List.sortBy (fun n -> n.Id), roundTripped.Numbering |> List.sortBy (fun n -> n.Id))
    Assert.Equal<DocumentProtection option>(doc.Protection, roundTripped.Protection)
    assertXmlSchemaValid (XDocument(xml))

/// The property that makes committing `document.xml` to source control and diffing it
/// across commits actually meaningful: two `Document` values with the same content but
/// differently-ordered `Styles`/`Numbering`/`TableStyles` (as they'd naturally be if e.g.
/// two real .docx files from different producers happened to declare the same style/
/// numbering catalog in a different order - these are ID-referenced catalogs, so their own
/// list order carries no semantic meaning, unlike `Sections`' real document order) must
/// produce byte-identical `toDocument` output. Without this, a re-generated XML could show
/// a spurious diff (styles shuffled) with no real content change - the same property the
/// Excel sibling's own `` Xml.ofWorkbook produces deterministic, input-order-independent
/// output `` test proves for `DefinedNames`.
[<Fact>]
let ``Xml.toDocument produces deterministic, input-order-independent output for Styles, Numbering, and TableStyles`` () =
    let docA =
        minimalDocument
        |> withStyles [ BuiltInStyles.heading1; BuiltInStyles.normal ]
        |> withNumbering [ numberedListDef 2; bulletListDef 1 ]
        |> withTableStyles
            [ { TableStyleDefinition.Default with Id = "Beta"; Name = "Beta" }
              { TableStyleDefinition.Default with Id = "Alpha"; Name = "Alpha" } ]

    let docB =
        { docA with
            Styles = docA.Styles |> List.rev
            Numbering = docA.Numbering |> List.rev
            TableStyles = docA.TableStyles |> List.rev }

    // docA and docB are *not* structurally equal as F# values (their lists are in different
    // orders) - the property under test is that they still render to identical XML.
    Assert.Equal((Xml.toDocument docA).ToString(), (Xml.toDocument docB).ToString())
