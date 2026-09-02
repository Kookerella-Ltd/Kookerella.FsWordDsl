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
    Assert.Equal<StyleDefinition list>(doc.Styles, roundTripped.Styles)
    Assert.Equal<NumberingDefinition list>(doc.Numbering, roundTripped.Numbering)
    Assert.Equal<DocumentProtection option>(doc.Protection, roundTripped.Protection)
    assertXmlSchemaValid (XDocument(xml))
