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
    Assert.Equal<StyleDefinition list>(doc.Styles, roundTripped.Styles)
    Assert.Equal<NumberingDefinition list>(doc.Numbering, roundTripped.Numbering)
    Assert.Equal<DocumentProtection option>(doc.Protection, roundTripped.Protection)
    assertJsonSchemaValid json
