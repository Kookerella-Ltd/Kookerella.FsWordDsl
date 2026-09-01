# CLAUDE.md

Instructions for any Claude Code session working in this repo.

## Repo layout

This repo currently ships **the F# core only** - no C# wrapper and no MCP server yet
(compare `Kookerella.FsOpenXmlDsl`, the Excel analog this repo was built to mirror, which
has all three). Don't assume either exists; if you're asked to add one, that's a new
package under `src/`, following the same pattern `Kookerella.CsOpenXmlDsl`/
`Kookerella.FsOpenXmlDsl.Mcp` do in that repo - not an extension of what's here.

- `src/Kookerella.FsWordDsl` - the F# core: a typesafe DSL over the WordprocessingML
  schema, interpreted by `Interpreter/Writer.fs` and reversed by `Interpreter/Reader.fs`.
- `tests/Kookerella.FsWordDsl.Tests` - one scenario per feature under `Examples/`, each
  validated against the real OOXML schema and round-tripped exactly back through the DSL.
- `samples/Kookerella.FsWordDsl.Sample` - a small console app exercising the DSL end to
  end (build, save, reload).

## A real, hard-won gotcha specific to this SDK/F# combination

**Never pass a single already-constructed `OpenXmlElement` as the sole positional argument
to another element's constructor** (e.g. `Wordprocessing.Document(body)`,
`Wordprocessing.Run(someChild)`). F# resolves a one-argument call like that to the SDK's
`IEnumerable<OpenXmlElement>` constructor overload, not "wrap this one child" - because
every `OpenXmlCompositeElement` (even leaf-ish ones like `Break`/`TabChar`) implements that
interface over its own children. The result is either:
- a **silently empty parent** if the argument has no children of its own (e.g.
  `Run(Break(...))` produces a `<w:r/>` with the `Break` just dropped), or
- a **runtime `InvalidOperationException: "...is part of a tree"`** if the argument already
  has children (e.g. `Document(body)` once `body` has any paragraphs appended).

Two-argument-or-more constructor calls are unaffected (arity alone rules out the
`IEnumerable` overload) and are used freely throughout `Writer.fs`. The fix for the
single-argument case is always: construct empty, then `.AppendChild(child)`. See
`Writer.fs`'s own note at `Document`'s construction and `ImageWriter.fs`'s own note at the
top of `addImage` for two worked examples - if you add a new single-child construction
anywhere, follow the same pattern, and verify with a quick `dotnet fsi` reflection check
(construct it, check `Seq.length x.ChildElements`) rather than assuming it's fine because
the equivalent C# sample on Microsoft Learn looks like `new Parent(new Child(...))`.

A related, permanent one: **F# cannot alias a .NET *namespace* as a `module`**
(`module W = DocumentFormat.OpenXml.Wordprocessing` is a compile error, FS0965 - module
abbreviations only work for actual modules). `Writer.fs`/`Reader.fs`/`StyleRegistry.fs`/
`ImageWriter.fs`/`ImageReader.fs` instead do `open DocumentFormat.OpenXml` and reference
the SDK's types via the nested namespace's own short name (`Wordprocessing.Paragraph`,
`Drawing.Wordprocessing.Inline`, `Drawing.Pictures.Picture`) - this works because F#
resolves a nested namespace by its own trailing segment once an ancestor namespace is
`open`, the same way C#'s `using` does. Never `open DocumentFormat.OpenXml.Wordprocessing`
directly in these files - it would shadow this DSL's own natural type/case names
(`Paragraph`, `Table`, `Hyperlink`, `Bookmark`, `Comment`, ...), which is exactly what this
qualification scheme avoids.

## Adding a feature to the DSL - checklist

1. **Model** (`src/Kookerella.FsWordDsl`): add the type(s) to the relevant feature file
   (or `Model.fs` if it participates in the `Block`/`Inline` recursion - see that file's
   own note on why some types live there instead of a dedicated feature file).
2. **Builders.fs**: a smart constructor on `DocumentDsl` if it's the kind of thing callers
   build directly (mirrors `SheetDsl` in the Excel repo).
3. **`Interpreter/Writer.fs`**: DSL -> OOXML. Watch for the single-child-constructor gotcha
   above.
4. **`Interpreter/Reader.fs`**: OOXML -> DSL, the literal inverse of step 3.
5. **`Interpreter/CodeGen.fs`**: DSL -> F# source text - a new field needs a new
   `render*`/matching case here too, or `Document.generateScript` silently omits it.
6. **`Xml.fs` + `Xml.xsd`**: the XML surface needs both the translation code and a matching
   schema change - `XmlTests.fs`'s `assertXmlSchemaValid` (used by every scenario's
   `document.xml`, plus its own direct round-trip tests) is what catches the two drifting
   apart.
7. **`Json.fs` + `Json.schema.json`**: same idea for the JSON surface -
   `JsonTests.fs`'s `assertJsonSchemaValid` is the equivalent check.
8. **Tests**: a new `Examples/<ScenarioName>` scenario in `tests/Kookerella.FsWordDsl.
   Tests/Tests.fs` demonstrating the feature (add its name to the `Category=Slow` theory's
   `InlineData` list too, so the regenerated-script check covers it), or extend an
   existing scenario if it's a small addition to something already covered.
9. **`MAPPING.md`**: update "Modeled faithfully" (or add to "Known gaps" if the feature is
   only partially modeled).
10. **README.md**: the layout list and, if it's a significant feature, a worked example
    matching the style of the existing ones.

## Process discipline

- Never commit or push without the user explicitly saying so in the current turn - a prior
  approval doesn't carry forward to later, unrelated changes.
- Verify, don't assume: run the actual build/test before reporting something works. This
  repo's own history (in this session) is proof of why - several plausible-looking single-
  and double-argument OOXML SDK constructor calls silently produced empty or malformed
  elements, caught only by actually running `OpenXmlValidator` and a real round trip, not
  by the code compiling or "looking right" against a Microsoft Learn C# sample.

## Build

- `dotnet build` (from the repo root, using `Kookerella.FsWordDsl.slnx`, or per-project).
- Fast tests only: `dotnet test --filter "Category!=Slow"` (from `tests/
  Kookerella.FsWordDsl.Tests`).
- Slow tests (actually executes every generated `Examples/*/script.fsx` via `dotnet fsi`
  and diffs the result's DSL structure against the committed file): `dotnet test --filter
  "Category=Slow"` - run this at least once after any `Writer.fs`/`Reader.fs`/`CodeGen.fs`
  change, not just the fast suite, and after the fast suite has populated the `.fsx` files
  at least once (it does, every time it runs).
- Plain `dotnet test` (no filter) runs both groups.
- `dotnet run --project samples/Kookerella.FsWordDsl.Sample` - builds a small report,
  saves it, and reads it back, printing a summary.
