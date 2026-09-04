module JsonTests

open System.IO
open System.Text.Json.Nodes
open Json.Schema
open Xunit
open Kookerella.FsWordDsl

/// Validates `node` against `Json.schema.json` - test-suite only, same posture as Excel's
/// own `Json.fs` doc comment explains (no built-in JSON Schema validator in .NET the way
/// `System.Xml.Schema` exists for XML). Shared by `Tests.verifyScenarioNamed` (every
/// scenario's `document.json`) and this file's own direct round-trip tests.
let private schemaPath = Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "src", "Kookerella.FsWordDsl", "Json.schema.json")
let private schema = JsonSchema.FromText(File.ReadAllText(schemaPath))

let assertJsonSchemaValid (node: JsonNode) =
    let element = System.Text.Json.JsonSerializer.SerializeToElement(node)
    let result = schema.Evaluate(element, EvaluationOptions(OutputFormat = OutputFormat.List))
    Assert.True(result.IsValid, string result)

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
let ``Json round trips a minimal document`` () =
    let json = Json.toDocument minimalDocument
    let roundTripped = Json.ofDocument json
    Assert.Equal<Section list>(minimalDocument.Sections, roundTripped.Sections)
    assertJsonSchemaValid json

[<Fact>]
let ``Json round trips styles, numbering, and protection`` () =
    let doc =
        minimalDocument
        |> withStyles [ BuiltInStyles.normal; BuiltInStyles.heading1 ]
        |> withNumbering [ bulletListDef 1 ]
        |> withProtection { Edit = Some ReadOnlyRestriction; Password = None }

    let json = Json.toDocument doc
    let roundTripped = Json.ofDocument json
    // Sorted before comparing on both sides - toDocument itself sorts Styles/Numbering by
    // Id (see its own doc comment), so content, not input order, is what this asserts.
    Assert.Equal<StyleDefinition list>(doc.Styles |> List.sortBy (fun s -> s.Id), roundTripped.Styles |> List.sortBy (fun s -> s.Id))
    Assert.Equal<NumberingDefinition list>(doc.Numbering |> List.sortBy (fun n -> n.Id), roundTripped.Numbering |> List.sortBy (fun n -> n.Id))
    Assert.Equal<DocumentProtection option>(doc.Protection, roundTripped.Protection)
    assertJsonSchemaValid json

/// The JSON equivalent of `XmlTests`' own `` Xml.toDocument produces deterministic,
/// input-order-independent output for Styles, Numbering, and TableStyles `` - see that
/// test's own doc comment for why this matters.
[<Fact>]
let ``Json.toDocument produces deterministic, input-order-independent output for Styles, Numbering, and TableStyles`` () =
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

    Assert.Equal((Json.toDocument docA).ToJsonString(), (Json.toDocument docB).ToJsonString())
