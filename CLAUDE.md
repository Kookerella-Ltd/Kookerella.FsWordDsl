# CLAUDE.md

Instructions for any Claude Code session working in this repo.

## Repo layout

This repo ships **the F# core plus a fluent C# wrapper** - no MCP server yet (compare
`Kookerella.FsOpenXmlDsl`, the Excel analog this repo was built to mirror, which has all
three). Don't assume the MCP server exists; if you're asked to add one, that's a new
package under `src/`, following the pattern `Kookerella.FsOpenXmlDsl.Mcp` does in that repo
- not an extension of what's here.

- `src/Kookerella.FsWordDsl` - the F# core: a typesafe DSL over the WordprocessingML
  schema, interpreted by `Interpreter/Writer.fs` and reversed by `Interpreter/Reader.fs`.
- `src/Kookerella.CsWordDsl` - an idiomatic, immutable, fluent C# wrapper over the F# core
  (`DocumentConverter.cs` does the two-way translation; `DocumentIO.cs` is the one place it
  touches I/O; `CsCodeGen.cs` is its own C#-source-text decompiler, the C# analog of
  `Interpreter/CodeGen.fs`) - see that project's own `DocumentConverter.cs` doc comment for
  the F#-compiled-shape gotchas this needed (DU cases as `New<Case>` static factories/
  singleton properties, case field names keeping their F#-source lowercase casing unlike a
  plain record's PascalCase properties, tuples as `System.Tuple`, not `ValueTuple`).
- `tests/Kookerella.FsWordDsl.Tests` - one scenario per feature under `Examples/`, each
  validated against the real OOXML schema and round-tripped exactly back through the DSL.
- `tests/Kookerella.CsWordDsl.Tests` - `DriftGuardTests.cs` (a reflection-based tripwire
  comparing F# DU case counts against their C# mirrors - see its own doc comment),
  `DocumentTests.cs` (targeted round-trip assertions per feature - not whole-`Document`
  equality, since `IReadOnlyList<T>` properties don't get deep structural equality for
  free), `ExampleTests.cs` (reloads the F# suite's own `Examples/*/output.docx` fixtures
  rather than re-authoring every scenario), `CsCodeGenTests.cs` (actually executes a
  generated file via `dotnet run --file`, the C# analog of the F# suite's `Category=Slow`
  `dotnet fsi` group).
- `samples/Kookerella.FsWordDsl.Sample` - a small console app exercising the F# DSL end to
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
11. **`src/Kookerella.CsWordDsl`**: add/extend the matching C# type(s) (same file-per-type
    convention the existing files use), then wire both directions into
    `DocumentConverter.cs` and the rendering into `CsCodeGen.cs`. `tests/
    Kookerella.CsWordDsl.Tests/DriftGuardTests.cs` will fail loudly if a new F# DU case
    doesn't get a matching C# case - that's the tripwire catching exactly this omission,
    not a test to silence by adding to its `KnownGaps` unless the omission is genuinely
    deliberate and documented.

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
- `dotnet test tests/Kookerella.CsWordDsl.Tests` - the C# wrapper's own suite (no
  fast/slow split; `CsCodeGenTests.cs` shells out to `dotnet run --file` itself, so a
  single `dotnet test` run already covers the C# analog of the F# suite's slow group).
