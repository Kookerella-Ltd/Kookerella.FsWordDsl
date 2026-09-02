# Kookerella.FsWordDsl.Mcp

<!-- mcp-name: io.github.MarkNicholls/fsworddsl-mcp -->

An [MCP](https://modelcontextprotocol.io) (Model Context Protocol) server that exposes
`Kookerella.FsWordDsl`'s Word read/write/code-generation capabilities as tools any
MCP-compatible AI agent can call directly - build a document, read one back, or regenerate
its F#, C#, XML, or JSON representation - without writing any code itself.

**Most Word libraries only go one direction**: build a document from scratch, or mutate an
existing one, through an imperative object model. This one also goes the other way - read
any existing `.docx`/`.docm` and hand back idiomatic, runnable F# or C# source (or plain
XML or JSON, each against a real schema) that rebuilds an equivalent file, and build a new
one from that XML/JSON directly. A decompiler for Word documents, not just a writer. That's
available three ways from this one binary: as MCP tools (`generate_fsharp_script`/
`generate_csharp_script`/`generate_xml`/`create_document_from_xml`/`generate_json`/
`create_document_from_json`, plus `generate_xml_schema`/`generate_json_schema` for the
schemas themselves) for an AI agent, as a plain CLI (`fsworddsl-mcp convert`/`build`) for
anyone who isn't going through an MCP client, and as direct library calls
(`Document.generateScript`/`CsCodeGen.Generate`/`Xml.toDocument`/`Xml.ofDocument`/
`Json.toDocument`/`Json.ofDocument`) for either to call themselves.

The XML/JSON directions each have two concrete uses beyond code generation: **build a
`.docx` from XML/JSON a transform engine already produces** (an XSLT pipeline, or any
templating/generation script, can target Word with no code at all), and **convert an
existing `.docx` to XML/JSON for version control** - `.docx` is a binary ZIP, so `git diff`
on one is useless, but `generate_xml`/`generate_json`'s output is plain, ordered text, so a
real content change produces a legible diff instead of an opaque binary one.

This runs **locally**, as a subprocess your MCP client launches over stdio - there's no
hosted service, no network address, and no account to sign up for. It's distributed as a
[.NET tool](https://learn.microsoft.com/dotnet/core/tools/global-tools) for exactly that
reason: an MCP client just needs a command it can run, the same way it'd run any other CLI.

## Install

```bash
dotnet tool install -g Kookerella.FsWordDsl.Mcp
```

This installs the `fsworddsl-mcp` command onto your PATH.

## Configure your MCP client

Point your client at the installed command. For example, in a client that reads a JSON
config with a `mcpServers` map:

```json
{
  "mcpServers": {
    "fsworddsl": {
      "command": "fsworddsl-mcp"
    }
  }
}
```

## Command-line usage

The same binary also works as a plain CLI, for converting a file without an MCP client at
all - `fsworddsl-mcp` with no arguments starts the MCP server (as above); with a `convert`
or `build` first argument it runs once and exits:

```bash
fsworddsl-mcp convert report.docx --lang csharp
```

Prints the equivalent C# source to stdout. Options:

- `--lang`/`-l` (required) — `fsharp`, `csharp`, `xml`, or `json`.
- `-o`/`--output <file>` — write the result to a file instead of stdout.
- `--rebuild-as <name.docx>` — `fsharp`/`csharp` only, the filename the *generated script
  itself* saves its rebuilt document to when run (default `output.docx`). Ignored for
  `--lang xml`/`json`, which have no script to embed a save path into.

```bash
fsworddsl-mcp convert report.docx --lang fsharp -o report.fsx --rebuild-as rebuilt.docx
dotnet fsi report.fsx
```

`build` is the inverse of `convert --lang xml`/`convert --lang json` - it takes XML matching
`Xml.xsd` or JSON matching `Json.schema.json` and produces a `.docx` directly, for a caller
(e.g. an XSLT pipeline, or a plain JSON-emitting script) that already produces data that way
and wants to reach Word without writing any code. Which format `build` reads is inferred
from the input file's own extension (`.xml` or `.json`):

```bash
fsworddsl-mcp convert report.docx --lang xml -o report.xml   # .docx -> XML
fsworddsl-mcp build report.xml rebuilt.docx                  # XML -> .docx

fsworddsl-mcp convert report.docx --lang json -o report.json # .docx -> JSON
fsworddsl-mcp build report.json rebuilt.docx                 # JSON -> .docx
```

## Tools

- **`create_document(path, paragraphs)`** — creates a new `.docx` from a simple list of
  paragraph texts. Each string becomes one plain paragraph, in order - no formatting,
  styles, tables, or images.
- **`read_document(path)`** — reads an existing `.docx`/`.docm` and returns its paragraphs
  as a JSON array of plain strings, using the same convention as `create_document`'s input
  (so a value round-trips through both tools unchanged). Table content and anything outside
  the plain-paragraph model is skipped.
- **`generate_fsharp_script(path, outputFileName)`** — reads an existing document and
  returns a self-contained F# script that rebuilds an equivalent file when run via
  `dotnet fsi`, using the full `Kookerella.FsWordDsl` API.
- **`generate_csharp_script(path, outputFileName)`** — the C# equivalent, for a caller who
  wants pasteable/runnable C# rather than F#. Reads an existing document through
  `Kookerella.CsWordDsl` (the idiomatic C# wrapper - full feature parity with the F# core)
  and returns a self-contained `.cs` file targeting .NET 10's file-based apps feature:
  `dotnet run <file>.cs`, no `.csproj` needed. References the published
  `Kookerella.CsWordDsl` NuGet package (via a `#:package` directive pinned to the version
  this server was built against), so the result runs on any machine with the .NET 10 SDK,
  not just this one.
- **`generate_xml(path)`** — a plain-data alternative to the two `generate_*_script` tools:
  reads an existing document and returns it as XML, validated against
  `Kookerella.FsWordDsl`'s own embedded schema (`Xml.xsd`). No `outputFileName` parameter,
  since the result is data, not a runnable script with a save path to embed.
- **`create_document_from_xml(xml, path)`** — the inverse of `generate_xml`: builds a new
  document from XML matching `Xml.xsd` and saves it to `path`. Unlike `create_document`,
  this isn't limited to plain paragraph text - named styles, tables, images, and every
  other modeled feature can be expressed in the XML, since it goes through the same schema
  `generate_xml` produces.
- **`generate_json(path)`** — the JSON equivalent of `generate_xml`: reads an existing
  document and returns it as JSON. Unlike `generate_xml`, there's no runtime JSON Schema
  validation built into the core library itself (see `generate_json_schema` below for why),
  but the documented shape is still retrievable.
- **`create_document_from_json(json, path)`** — the inverse of `generate_json`: builds a
  new document from JSON matching the shape `generate_json` produces and saves it to
  `path`. Same relationship to `create_document` as `create_document_from_xml` has - not
  limited to plain paragraph text.
- **`generate_xml_schema()`** — returns the raw XSD (`Xml.xsd`) that `generate_xml`/
  `create_document_from_xml` conform to, so a caller authoring XML by hand or by transform
  (e.g. an XSLT stylesheet) can get real schema validation/autocomplete in their own editor
  or pipeline instead of reverse-engineering the shape from an example.
- **`generate_json_schema()`** — the JSON equivalent: returns the raw JSON Schema
  (`Json.schema.json`) that `generate_json`/`create_document_from_json` conform to. This
  schema isn't validated against at runtime by the core library the way `Xml.xsd` is - JSON
  Schema has no .NET-built-in equivalent to `System.Xml.Schema`, so wiring that up there
  would mean adding a runtime dependency (`JsonSchema.Net`) to every consumer of the core
  library just for this. It's bundled in this Mcp tool specifically, purely to hand back on
  request.

## Scope

`create_document`/`read_document` are a deliberately narrow first pass over the library,
not the whole thing: plain paragraph text only, one flat paragraph per string, no run
styling, named styles, tables, images, headers/footers, comments, track changes, or content
controls. An agent that needs those should reference the library directly, or use one of
the other six document-decompiling tools on a file that already has them to see it
represented as source or data (`generate_xml_schema`/`generate_json_schema` are a separate
pair - they don't take a document at all, just hand back the schema itself). All six cover
the full section/paragraph-level feature set: `generate_fsharp_script` the full F# core,
`generate_csharp_script`/`generate_xml`/`create_document_from_xml`/`generate_json`/
`create_document_from_json` everything `Kookerella.CsWordDsl`/`Xml.fs`/`Json.fs` model,
which are the same feature set as the F# core. See the main project's
[MAPPING.md](https://github.com/Kookerella-Ltd/Kookerella.FsWordDsl/blob/master/MAPPING.md)
for the full picture of what the underlying library does and doesn't model.
