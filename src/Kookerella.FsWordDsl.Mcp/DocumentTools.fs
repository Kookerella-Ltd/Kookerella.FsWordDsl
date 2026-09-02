namespace Kookerella.FsWordDsl.Mcp

open System
open System.ComponentModel
open System.IO
open System.Text.Json
open System.Text.Json.Nodes
open System.Xml.Linq
open ModelContextProtocol.Server
open Kookerella.FsWordDsl

/// The MCP tool surface over `Kookerella.FsWordDsl`. `create_document`/`read_document` are
/// deliberately narrow - plain paragraph text only, one string per paragraph, no
/// styling/tables/images/etc. - the same "honest, bounded MVP, documented gap" scoping the
/// Excel sibling's own `WorkbookTools`/`create_workbook`/`read_workbook` uses. The other six
/// tools (`generate_fsharp_script`/`generate_csharp_script`/`generate_xml`/
/// `create_document_from_xml`/`generate_json`/`create_document_from_json`) aren't limited
/// that way - they cover the full section/paragraph-level feature set, just via generated
/// source/XML/JSON rather than a plain paragraph list. `generate_xml_schema`/
/// `generate_json_schema` are a separate pair again - they take no document at all, they
/// just hand back the schema those two directions conform to.
[<McpServerToolType>]
type DocumentTools =

    /// One flat paragraph per input string, no direct/named formatting - the same "plain
    /// text, no styling" narrowness `WorkbookTools.ToWorksheet`'s own cell grid uses on the
    /// Excel side.
    static member private ToParagraphBlock(text: string) : Block =
        ParagraphBlock
            { Inlines = [ Run(text, None, None) ]
              StyleId = None
              Format = None
              Numbering = None
              MarkRevision = None }

    /// The inverse of `ToParagraphBlock` - concatenates a paragraph's own `Run` inlines
    /// back into one plain string, dropping anything else (images, hyperlinks, bookmarks,
    /// content controls, ...) that doesn't fit the narrow "plain text" shape this tool
    /// exposes. Non-paragraph blocks (tables, content control blocks) are skipped entirely
    /// rather than partially rendered.
    static member private ParagraphText(p: Paragraph) : string =
        p.Inlines
        |> List.choose (function
            | Run(text, _, _) -> Some text
            | _ -> None)
        |> String.concat ""

    [<McpServerTool(Name = "create_document")>]
    [<Description(
        "Creates a new Word document (.docx) from a simple list of paragraph texts and saves it to disk. \
         Each string becomes one plain paragraph, in order - no formatting, styles, tables, or images in this \
         version. Does not support run styling, tables, images, headers/footers, or track changes - reference \
         the Kookerella.FsWordDsl library directly for those."
    )>]
    static member CreateDocument
        (
            [<Description("Output file path, e.g. \"C:\\reports\\memo.docx\". The directory must already exist.")>] path: string,
            [<Description("The paragraphs to create, in order. Each element becomes one plain paragraph.")>] paragraphs: string[]
        ) : string =
        let body = paragraphs |> Array.toList |> List.map DocumentTools.ToParagraphBlock

        let doc: Document =
            { Sections = [ { Body = body; Properties = SectionProperties.Default } ]
              Styles = BuiltInStyles.all
              Numbering = []
              Protection = None
              VbaProject = None
              Properties = DocumentProperties.Default
              TableStyles = [] }

        Document.save path doc
        sprintf "Wrote %s (%d paragraph%s)." path paragraphs.Length (if paragraphs.Length = 1 then "" else "s")

    [<McpServerTool(Name = "read_document")>]
    [<Description(
        "Reads an existing Word document (.docx/.docm) and returns its paragraphs as a JSON array of plain \
         strings, one per paragraph, matching create_document's own input convention. Features outside the \
         plain-paragraph model (run styling, tables, images, headers/footers, comments, track changes, content \
         controls, etc.) are not included in this output - see MAPPING.md in the main library repo for the \
         full list of what round-trips. Table content and non-paragraph blocks are skipped entirely."
    )>]
    static member ReadDocument([<Description("Path to an existing .docx or .docm file.")>] path: string) : string =
        let doc = Document.load path

        let texts =
            doc.Sections
            |> List.collect (fun s -> s.Body)
            |> List.choose (function
                | ParagraphBlock p -> Some(DocumentTools.ParagraphText p)
                | _ -> None)
            |> List.toArray

        JsonSerializer.Serialize(texts, JsonSerializerOptions(WriteIndented = true))

    [<McpServerTool(Name = "generate_fsharp_script")>]
    [<Description(
        "Reads an existing Word document and returns a self-contained F# script (using Kookerella.FsWordDsl) \
         that rebuilds an equivalent file when run via `dotnet fsi`. Useful for explaining how a file is \
         structured, or as a starting point for a caller who wants the library's full feature set (named \
         styles, tables, images, headers/footers, comments, track changes, content controls, etc.) beyond what \
         create_document exposes."
    )>]
    static member GenerateFSharpScript
        (
            [<Description("Path to an existing .docx/.docm file to reverse-engineer into F# source.")>] path: string,
            [<Description("The output filename the generated script should save its rebuilt file to, e.g. \"output.docx\".")>] outputFileName: string
        ) : string =
        let doc = Document.load path

        // Portable across machines - #r "nuget: ..." pulls the package via FSI's own NuGet
        // resolution (DocumentFormat.OpenXml comes in transitively, the same way it does for
        // any normal project reference), matching whatever version of the core this server
        // itself was built against. A raw #r "<dll path>" only works on the machine this
        // server happens to be installed on - see generate_csharp_script's own #:package
        // for the C# equivalent of this fix.
        let packageVersion: Version = typeof<Document>.Assembly.GetName().Version

        let referenceLines =
            [ sprintf "#r \"nuget: Kookerella.FsWordDsl, %d.%d.%d\"" packageVersion.Major packageVersion.Minor packageVersion.Build ]

        Document.generateScript referenceLines outputFileName doc

    [<McpServerTool(Name = "generate_csharp_script")>]
    [<Description(
        "Reads an existing Word document and returns a self-contained C# file (using Kookerella.CsWordDsl) that \
         rebuilds an equivalent file when run via `dotnet run <file>.cs` (.NET 10's file-based apps feature - \
         no .csproj needed). The C# equivalent of generate_fsharp_script, for a caller who wants pasteable/ \
         runnable C# rather than F# - useful for explaining how a file is structured, or as a starting point \
         for the wrapper's fluent API (named styles, tables, images, headers/footers, comments, track changes, \
         content controls, protection, etc. - Kookerella.CsWordDsl covers the same feature set as \
         generate_fsharp_script's own Kookerella.FsWordDsl) beyond what create_document exposes."
    )>]
    static member GenerateCSharpScript
        (
            [<Description("Path to an existing .docx/.docm file to reverse-engineer into C# source.")>] path: string,
            [<Description("The output filename the generated script should save its rebuilt file to, e.g. \"output.docx\".")>] outputFileName: string
        ) : string =
        let doc = Kookerella.CsWordDsl.DocumentIO.Load(path)

        // Unlike generate_fsharp_script's `#r` (a raw, machine-specific DLL path - .NET
        // file-based apps don't support `#r` at all, only `#:package`/`#:project`), this
        // points at the published NuGet package matching whatever version of the wrapper
        // this server itself was built against - portable to any machine with the .NET 10
        // SDK, not just this one.
        let packageVersion: Version = typeof<Kookerella.CsWordDsl.Document>.Assembly.GetName().Version

        let referenceLines =
            [| sprintf "#:package Kookerella.CsWordDsl@%d.%d.%d" packageVersion.Major packageVersion.Minor packageVersion.Build |]

        Kookerella.CsWordDsl.CsCodeGen.Generate(referenceLines, outputFileName, doc)

    [<McpServerTool(Name = "generate_xml")>]
    [<Description(
        "Reads an existing Word document and returns it as XML, validated against Kookerella.FsWordDsl's own \
         embedded schema (Xml.xsd). A plain-data alternative to generate_fsharp_script/generate_csharp_script \
         for a caller who wants to inspect, transform (e.g. via XSLT), or archive a document's structure \
         without any F#/C# source involved - and without any .NET runtime on the caller's side either: the \
         .NET work happens inside this server, so Python, JavaScript, or any other language can call this tool \
         directly, and a human with no MCP client at all can get the same result via \
         `fsworddsl-mcp convert <file> --lang xml` from a plain shell. Unlike those two, this returns data, not \
         a runnable script, so there is no output-filename parameter to control what a rebuild saves as. \
         Covers the same section/paragraph-level feature set generate_fsharp_script does."
    )>]
    static member GenerateXml([<Description("Path to an existing .docx/.docm file to convert to XML.")>] path: string) : string =
        let doc = Document.load path
        (Xml.toDocument doc).ToString()

    [<McpServerTool(Name = "create_document_from_xml")>]
    [<Description(
        "Builds a new Word document from XML matching Kookerella.FsWordDsl's own embedded schema (Xml.xsd) and \
         saves it to disk - the inverse of generate_xml. The natural target for a caller that already produces \
         data as XML (e.g. an XSLT pipeline generating a report) and wants to reach Word without learning the \
         OOXML schema, this library's own F#/C# API, or needing .NET installed at all - like generate_xml, the \
         .NET work happens inside this server, so any language can call it directly, and a human with no MCP \
         client at all can get the same result via `fsworddsl-mcp build` from a plain shell. Covers the same \
         section/paragraph-level feature set generate_xml does; unlike create_document, this isn't limited to \
         plain paragraph text - named styles, tables, images, and every other modeled feature can be expressed \
         in the XML."
    )>]
    static member CreateDocumentFromXml
        (
            [<Description("The document XML content - a <document> root element matching Xml.xsd.")>] xml: string,
            [<Description("Output file path, e.g. \"C:\\reports\\memo.docx\". The directory must already exist.")>] path: string
        ) : string =
        let doc = XElement.Parse(xml) |> Xml.ofDocument
        Document.save path doc
        let paragraphCount = doc.Sections |> List.sumBy (fun s -> s.Body |> List.length)
        sprintf "Wrote %s (%d section%s, %d block%s)." path doc.Sections.Length (if doc.Sections.Length = 1 then "" else "s") paragraphCount (if paragraphCount = 1 then "" else "s")

    [<McpServerTool(Name = "generate_json")>]
    [<Description(
        "Reads an existing Word document and returns it as JSON. The JSON-side equivalent of generate_xml, for \
         a caller whose tooling speaks JSON rather than XML - same use cases (inspect, transform, or archive a \
         document's structure without any F#/C# source, or any .NET runtime at all, on the caller's side) and \
         the same section/paragraph-level feature set. Usable directly from Python, JavaScript, or any other \
         language, and a human with no MCP client at all can get the same result via \
         `fsworddsl-mcp convert <file> --lang json` from a plain shell. Unlike generate_xml, there's no runtime \
         JSON Schema validation built into the core library itself (see generate_json_schema's own doc string \
         for why) - but generate_json_schema still returns the documented shape this produces."
    )>]
    static member GenerateJson([<Description("Path to an existing .docx/.docm file to convert to JSON.")>] path: string) : string =
        let doc = Document.load path
        (Json.toDocument doc).ToJsonString(JsonSerializerOptions(WriteIndented = true))

    [<McpServerTool(Name = "create_document_from_json")>]
    [<Description(
        "Builds a new Word document from JSON matching the shape generate_json produces (see the main library \
         repo's Json.schema.json) and saves it to disk - the inverse of generate_json. The JSON-side equivalent \
         of create_document_from_xml, for a caller that already produces data as JSON and wants to reach Word \
         without learning the OOXML schema, this library's own F#/C# API, or needing .NET installed at all - \
         the .NET work happens inside this server, so any language can call it directly, and a human with no \
         MCP client at all can get the same result via `fsworddsl-mcp build` from a plain shell. Covers the \
         same section/paragraph-level feature set generate_json does; unlike create_document, this isn't \
         limited to plain paragraph text - named styles, tables, images, and every other modeled feature can \
         be expressed in the JSON."
    )>]
    static member CreateDocumentFromJson
        (
            [<Description("The document JSON content - an object matching Json.schema.json's root shape.")>] json: string,
            [<Description("Output file path, e.g. \"C:\\reports\\memo.docx\". The directory must already exist.")>] path: string
        ) : string =
        let doc = JsonNode.Parse(json).AsObject() |> Json.ofDocument
        Document.save path doc
        sprintf "Wrote %s (%d section%s)." path doc.Sections.Length (if doc.Sections.Length = 1 then "" else "s")

    [<McpServerTool(Name = "generate_xml_schema")>]
    [<Description(
        "Returns the raw XSD (Xml.xsd) that generate_xml's output and create_document_from_xml's input both \
         conform to. Meant for a caller authoring XML by hand or by transform (e.g. an XSLT stylesheet) who \
         wants real schema validation/autocomplete in their own editor or pipeline, rather than reverse- \
         engineering the shape from a generate_xml example."
    )>]
    static member GenerateXmlSchema() : string =
        let assembly = typeof<Document>.Assembly
        use stream = assembly.GetManifestResourceStream("Kookerella.FsWordDsl.Xml.xsd")
        use reader = new StreamReader(stream)
        reader.ReadToEnd()

    [<McpServerTool(Name = "generate_json_schema")>]
    [<Description(
        "Returns the raw JSON Schema (Json.schema.json) that generate_json's output and \
         create_document_from_json's input both conform to. Meant for a caller authoring JSON by hand or by a \
         generation script who wants real schema validation/autocomplete in their own editor or pipeline, \
         rather than reverse-engineering the shape from a generate_json example. This schema isn't validated \
         against at runtime by the core library itself the way Xml.xsd is (JSON Schema has no .NET-built-in \
         equivalent to System.Xml.Schema, so wiring that up would mean adding a runtime dependency - \
         JsonSchema.Net - to every consumer of the core library just for this) - it's bundled here, in the Mcp \
         tool specifically, purely to hand back on request."
    )>]
    static member GenerateJsonSchema() : string =
        let assembly = Reflection.Assembly.GetExecutingAssembly()
        use stream = assembly.GetManifestResourceStream("Kookerella.FsWordDsl.Mcp.Json.schema.json")
        use reader = new StreamReader(stream)
        reader.ReadToEnd()
