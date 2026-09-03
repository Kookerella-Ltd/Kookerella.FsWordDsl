# CLAUDE.md

Instructions for any Claude Code session working in this repo.

## Repo layout

This repo ships **the F# core, a fluent C# wrapper, and an MCP server** - full parity with
`Kookerella.FsOpenXmlDsl`, the Excel analog this repo was built to mirror.

- `src/Kookerella.FsWordDsl` - the F# core: a typesafe DSL over the WordprocessingML
  schema, interpreted by `Interpreter/Writer.fs` and reversed by `Interpreter/Reader.fs`.
- `src/Kookerella.CsWordDsl` - an idiomatic, immutable, fluent C# wrapper over the F# core
  (`DocumentConverter.cs` does the two-way translation; `DocumentIO.cs` is the one place it
  touches I/O; `CsCodeGen.cs` is its own C#-source-text decompiler, the C# analog of
  `Interpreter/CodeGen.fs`) - see that project's own `DocumentConverter.cs` doc comment for
  the F#-compiled-shape gotchas this needed (DU cases as `New<Case>` static factories/
  singleton properties, case field names keeping their F#-source lowercase casing unlike a
  plain record's PascalCase properties, tuples as `System.Tuple`, not `ValueTuple`).
- `src/Kookerella.FsWordDsl.Mcp` - an MCP (Model Context Protocol) server exposing the F#
  core's read/write/decompile capabilities as tools (`create_document`/`read_document` -
  deliberately narrow, plain-paragraph-text only; `generate_fsharp_script`/
  `generate_csharp_script`/`generate_xml`/`create_document_from_xml`/`generate_json`/
  `create_document_from_json` - full feature parity; `generate_xml_schema`/
  `generate_json_schema`). The same binary is also a plain CLI (`fsworddsl-mcp
  convert`/`build`) for callers not going through an MCP client - `Program.fs`'s
  `[<EntryPoint>]` dispatches on `argv`'s first token (`"convert"`/`"build"`/otherwise ->
  MCP stdio server). `DocumentTools.fs` is the tool surface, one `[<McpServerTool>]`-tagged
  static member per tool, mirroring the Excel sibling's own `WorkbookTools.fs` file-for-file.
- `tests/Kookerella.FsWordDsl.Tests` - one scenario per feature under `Examples/`, each
  validated against the real OOXML schema and round-tripped exactly back through the DSL.
- `tests/Kookerella.CsWordDsl.Tests` - `DriftGuardTests.cs` (a reflection-based tripwire
  comparing F# DU case counts against their C# mirrors - see its own doc comment),
  `DocumentTests.cs` (targeted round-trip assertions per feature - not whole-`Document`
  equality, since `IReadOnlyList<T>` properties don't get deep structural equality for
  free), `ExampleTests.cs` (reloads the F# suite's own `Examples/*/output.docx` fixtures
  rather than re-authoring every scenario), `CsCodeGenTests.cs` (actually executes a
  generated file via `dotnet run --file`, the C# analog of the F# suite's `Category=Slow`
  `dotnet fsi` group). There's no dedicated test project for the Mcp server, matching the
  Excel sibling's own posture - verify it by hand (`dotnet build` then either the CLI
  `convert`/`build` subcommands directly, or a real JSON-RPC handshake over stdio) after
  any change, the same "run the actual thing" discipline this file's own Process
  discipline section already asks for everywhere else.
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
- Never let a secret (API key, token) reach a command wrapper that logs its own full
  argument list - FAKE's `DotNet.exec` does this on every invocation, success or failure.
  Shell out via a raw process call with a redacted log line instead (see `build.fsx`'s
  `push` function for the pattern). Verify this by actually running the command and
  grepping captured output for the secret, not by reading the code and assuming it's safe.

## Release

This repo mirrors its Excel sibling's (`Kookerella.FsOpenXmlDsl`) build/release tooling
file-for-file - `build.fsx` (FAKE) drives Build/Test/Pack/Push/PublishAll plus the two
standalone-distribution targets below, with the same reasoning behind every design choice
documented inline in that file's own comments.

1. Bump `<Version>` by hand in whichever project(s) changed (a semver judgment call, not
   automated).
2. Run `dotnet fake run build.fsx -t PublishAll` (runs the full test gate first). **Use
   `PublishAll`, never `Push<Core|Wrapper|Mcp>` individually, and never a manual `dotnet
   pack`/`dotnet nuget push`** - `push` (in `build.fsx`) is deliberately a no-op for
   whichever package(s) didn't change this release, specifically so `PublishAll` is always
   safe and always the right thing to run. `PublishAll`'s own action then checks nuget.org's
   live state and fails loudly if either the C# wrapper's published NuGet dependency floor
   on the F# core, or the Mcp tool's *bundled* core/wrapper DLL versions (it's a
   self-contained `dotnet tool`, no nuspec dependency floor to check instead), still don't
   match the latest published core/wrapper versions - independently invocable any time via
   `dotnet fake run build.fsx -t VerifyDependencyFreshness --single-target`. **This isn't
   hypothetical**: standing this tooling up for the first time and running
   `VerifyDependencyFreshness` immediately caught a real case - the same-day
   `Kookerella.CsWordDsl` 0.1.1 README fix left the already-published
   `Kookerella.FsWordDsl.Mcp` 0.1.0 bundling a stale `Kookerella.CsWordDsl.dll` 0.1.0.0.
   NuGet indexing typically takes 5-20 minutes after a successful push before the new
   version resolves anywhere (search, flatcontainer index, `dotnet restore`) - don't assume
   a push failed just because it isn't visible yet.
3. **MCP Registry sync**, only if the Mcp package's version changed: `mcp-publisher login
   github` (interactive GitHub device-flow - this needs the user, it can't be scripted or
   run non-interactively) immediately followed by `mcp-publisher publish` from
   `src/Kookerella.FsWordDsl.Mcp/.mcp/` - the registry JWT is short-lived, so log in again
   right before publishing rather than reusing an older session. The registry publish will
   itself reject the request with a clear error if the NuGet version it references isn't
   indexed yet - that's the signal to wait, not a real failure.
4. **Standalone binaries / GitHub Release**, only when worth cutting a new platform build:
   `dotnet fake run build.fsx -t PackMcpSelfContained` (six self-contained single-file
   executables, one per `win-x64`/`win-arm64`/`linux-x64`/`linux-arm64`/`osx-x64`/
   `osx-arm64`) and/or `-t PackMcpMcpb` (the same six, repackaged as Claude Desktop's
   one-click `.mcpb` bundle format - needs `npm install -g @anthropic-ai/mcpb` on PATH
   first). Deliberately **not** chained into `PublishAll` - uploading these as GitHub
   Release assets is its own separate, explicitly-triggered step (`gh release create
   v<version> <assets...>`), the same reasoning as the MCP Registry sync needing a human
   login: this needs a human decision about when a new platform build is worth cutting, not
   something that should happen on every NuGet release automatically. Native AOT was tried
   and rejected empirically for this target, not just assumed impractical - see `build.fsx`'s
   own comment on `selfContainedRids` for why (F#'s `sprintf`/`printf`/`failwithf` machinery
   is reflection-based in a way Native AOT's trimmer can't resolve statically).

## Keep these in sync

A version bump or new feature can make any of these stale independently of the others -
useful as a quick scan regardless of what kind of change is in flight.

| File | What must stay accurate | Checked by |
|---|---|---|
| `README.md` (root) | Top summary, per-feature sections, Layout list | Nothing automated - read it |
| `src/Kookerella.CsWordDsl/README.md` | Feature list, `## Scope` (packed directly into the NuGet listing - verify by unpacking the built `.nupkg`, not just reading source) | Nothing automated |
| `src/Kookerella.FsWordDsl.Mcp/README.md` | Tool list, `## Scope`, CLI usage (packed directly into the NuGet listing, same caveat) | Nothing automated |
| `src/Kookerella.FsWordDsl/Kookerella.FsWordDsl.fsproj` `<Description>` | Matches the F# core's actual feature set | Nothing automated - it's NuGet metadata, not code |
| `src/Kookerella.CsWordDsl/Kookerella.CsWordDsl.csproj` `<Description>` | Matches the C# wrapper's actual feature set | Nothing automated |
| `src/Kookerella.FsWordDsl.Mcp/Kookerella.FsWordDsl.Mcp.fsproj` `<Description>` | Matches the Mcp server's actual tool/CLI surface | Nothing automated |
| `src/Kookerella.FsWordDsl.Mcp/.mcp/server.json` | `description` (≤100 chars, registry-enforced) **and** both `version` fields (top-level and `packages[].version`) match the `.fsproj`'s `<Version>` | The registry publish itself rejects a version it can't find on NuGet, and `VerifyDependencyFreshness`/`PublishAll`'s own action (`verifyServerJsonVersion`) checks both version fields against the `.fsproj` - but nothing checks `description` accuracy |
| `src/Kookerella.FsWordDsl.Mcp/DocumentTools.fs` | Every tool's `[<Description>]` text, and the `DocumentTools` type's own doc comment | Nothing automated - this is what an MCP client/agent actually reads, separate from any README |
| `src/Kookerella.FsWordDsl.Mcp/Dockerfile` | `COPY` list mirrors every `ProjectReference` in the `.fsproj` exactly | Nothing automated unless someone actually runs `docker build` |
| `src/Kookerella.FsWordDsl/Xml.xsd` | Matches what `Xml.fs`'s `ofDocument`/`toDocument` actually read and write | `assertXmlSchemaValid` inside `verifyScenarioNamed` - real, but only as strong as the scenarios that exist |
| Published `Kookerella.CsWordDsl`'s NuGet dependency floor on `Kookerella.FsWordDsl` | Must equal the F# core's latest *published* version, not just whatever's in the local `.fsproj` | `VerifyDependencyFreshness`/`PublishAll`'s own action in `build.fsx` (`verifyWrapperDependencyFloor`) - queries nuget.org's live state directly; only catches it if `PublishAll` (not a standalone `Push*`) is actually what ran |
| Published `Kookerella.FsWordDsl.Mcp`'s *bundled* `Kookerella.FsWordDsl.dll`/`Kookerella.CsWordDsl.dll` (it's a self-contained `dotnet tool`, no nuspec dependency floor to check instead) | Their assembly versions must equal the core's/wrapper's latest *published* versions | `VerifyDependencyFreshness`/`PublishAll`'s own action in `build.fsx` (`verifyMcpBundleFreshness`) - downloads the live `.nupkg` and reads the bundled DLLs' own version metadata; found live and stale the very first time this check was stood up in this repo (see Release step 2 above) |
| `src/Kookerella.FsWordDsl/Json.schema.json` | Matches what `Json.fs`'s `ofDocument`/`toDocument` actually read and write | `assertJsonSchemaValid`, called from every `JsonTests.fs` round trip, **and** from `verifyScenarioNamed` for every `Examples/` scenario (writes `document.json`, same as `document.xml`/`Xml.xsd`) - real, but only as strong as the scenarios that exist |
| `tests/Kookerella.CsWordDsl.Tests/DriftGuardTests.cs`'s enum/closed-hierarchy mirror lists | Lists every F# DU mirrored as a C# enum/closed hierarchy | Only guards types already registered with it - a brand-new type is invisible until added |
| `MAPPING.md` | What the F# core maps 1:1 vs. approximates vs. doesn't model | Nothing automated - only touch this for a new OOXML-level capability, not a wrapper-level one |

Three packages, three `.fsproj`/`.csproj` `<Version>` fields, one `server.json` with two
more copies of one of those three - a version bump in one place and not the others is the
single most common way this list goes stale. When in doubt, grep the whole repo for the old
version string before considering a bump finished.

## Build

- `dotnet tool restore` once (restores `fake-cli` from `.config/dotnet-tools.json`).
- `dotnet fake run build.fsx -t <Target>` is the primary way to build/test/release - see
  `build.fsx` for the full target list (`Clean`, `Restore`, `Build`, `TestFast`, `TestSlow`,
  `PackCore`/`PackWrapper`/`PackMcp`, `PushCore`/`PushWrapper`/`PushMcp`, `PublishAll`,
  `PackMcpSelfContained`, `PackMcpMcpb`) and the Release section above for the full
  publish sequence. The plain `dotnet`/CLI commands below still work directly for
  finer-grained iteration during day-to-day feature work.
- `dotnet build` (from the repo root, using `Kookerella.FsWordDsl.slnx`, or per-project).
- Fast tests only: `dotnet test --filter "Category!=Slow"` (from `tests/
  Kookerella.FsWordDsl.Tests`).
- Slow tests (actually executes every generated `Examples/*/script.fsx` via `dotnet fsi`
  and diffs the result's DSL structure against the committed file): `dotnet test --filter
  "Category=Slow"` - run this at least once after any `Writer.fs`/`Reader.fs`/`CodeGen.fs`
  change, not just the fast suite, and after the fast suite has populated the `.fsx` files
  at least once (it does, every time it runs). Running either suite rewrites the committed
  `Examples/*/output.docx`/`script.fsx` fixtures in place - OOXML zips aren't
  byte-deterministic even when structurally identical, so `git status` showing them as
  modified after a test run is expected noise, not a real change; `git checkout --
  tests/Kookerella.FsWordDsl.Tests/Examples/` before committing anything else.
- Plain `dotnet test` (no filter) runs both groups.
- `dotnet run --project samples/Kookerella.FsWordDsl.Sample` - builds a small report,
  saves it, and reads it back, printing a summary.
- `dotnet test tests/Kookerella.CsWordDsl.Tests` - the C# wrapper's own suite (no
  fast/slow split; `CsCodeGenTests.cs` shells out to `dotnet run --file` itself, so a
  single `dotnet test` run already covers the C# analog of the F# suite's slow group).
- The Mcp server has no dedicated test project - after any change there, at minimum
  `dotnet build src/Kookerella.FsWordDsl.Mcp` plus one CLI smoke test, e.g. `dotnet
  <path-to-Mcp.dll> convert <some Examples/*/output.docx> --lang json`. To verify the
  actual MCP stdio transport (not just the tool logic underneath it, which the CLI path
  already exercises), pipe a real JSON-RPC `initialize` + `tools/list` request pair in and
  confirm both get a `result` back - watch for Windows path strings needing `\\` (doubled)
  in the JSON, not a single `\`, which fails to parse rather than failing at the tool call.
