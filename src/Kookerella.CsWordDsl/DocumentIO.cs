using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Microsoft.FSharp.Collections;
using Fs = Kookerella.FsWordDsl;

namespace Kookerella.CsWordDsl;

/// <summary>
/// Every way this wrapper gets a <see cref="Document"/> in or out of some other
/// representation - every other type in this assembly is a pure, immutable value with no
/// I/O or conversion methods of its own (see <see cref="Document"/>'s own doc comment).
/// <see cref="Save(Document,string)"/>/<see cref="Load(string)"/> are the only members here
/// that actually touch a file or stream; <see cref="ToXml"/>/<see cref="FromXml"/>, <see
/// cref="ToJson"/>/<see cref="FromJson"/>, and <see cref="GenerateFSharpScript"/> are pure
/// string conversions, kept alongside them since a caller reaching for "get a Document from
/// X" naturally looks here first, the same way the F# core keeps <c>Document.save</c>/
/// <c>load</c>/<c>generateScript</c> together in one module rather than splitting by
/// side-effecting-ness. Every conversion here is a thin wrapper over the F# core's own
/// <c>Xml.fs</c>/<c>Json.fs</c>/<c>Document.generateScript</c> - this type does no
/// translation of its own beyond the F#&lt;-&gt;C# shape conversion <see
/// cref="DocumentConverter"/> already does for <see cref="Save(Document,string)"/>/<see
/// cref="Load(string)"/>, so a C# caller never needs to reference
/// <c>Kookerella.FsWordDsl</c> directly to reach any of these.
/// </summary>
public static class DocumentIO
{
    public static void Save(Document document, string path) =>
        Fs.Document.save(path, DocumentConverter.ToFSharp(document));

    public static void Save(Document document, Stream stream) =>
        Fs.Document.saveToStream(stream, DocumentConverter.ToFSharp(document));

    public static Document Load(string path) =>
        DocumentConverter.FromFSharp(Fs.Document.load(path));

    public static Document Load(Stream stream) =>
        DocumentConverter.FromFSharp(Fs.Document.loadFromStream(stream));

    /// <summary>Renders <paramref name="document"/> as XML matching the F# core's own
    /// embedded schema (<c>Xml.xsd</c>) - the C# entry point to the same surface
    /// <c>generate_xml</c> exposes as an MCP tool.</summary>
    public static string ToXml(Document document) =>
        Fs.Xml.toDocument(DocumentConverter.ToFSharp(document)).ToString();

    /// <summary>The inverse of <see cref="ToXml"/> - builds a <see cref="Document"/> from
    /// XML matching <c>Xml.xsd</c>'s <c>&lt;document&gt;</c> root element.</summary>
    public static Document FromXml(string xml) =>
        DocumentConverter.FromFSharp(Fs.Xml.ofDocument(XElement.Parse(xml)));

    /// <summary>Renders <paramref name="document"/> as JSON matching the shape
    /// <c>Json.schema.json</c> documents (see that schema's own comment for why it isn't
    /// validated against at runtime by the F# core the way <c>Xml.xsd</c> is) - the C#
    /// entry point to the same surface <c>generate_json</c> exposes as an MCP tool.
    /// </summary>
    public static string ToJson(Document document) =>
        Fs.Json.toDocument(DocumentConverter.ToFSharp(document)).ToJsonString(new JsonSerializerOptions { WriteIndented = true });

    /// <summary>The inverse of <see cref="ToJson"/> - builds a <see cref="Document"/> from
    /// JSON matching <c>Json.schema.json</c>'s root shape.</summary>
    public static Document FromJson(string json) =>
        DocumentConverter.FromFSharp(Fs.Json.ofDocument(JsonNode.Parse(json)!.AsObject()));

    /// <summary>Renders <paramref name="document"/> as a self-contained F# script that,
    /// when run via <c>dotnet fsi</c>, rebuilds an equivalent file at <paramref
    /// name="outputFileName"/> - the C# entry point to the same capability <see
    /// cref="CsCodeGen.Generate"/> offers for C#, and that <c>generate_fsharp_script</c>
    /// exposes as an MCP tool. <paramref name="referenceLines"/> are whatever raw
    /// <c>#r</c> directives the script needs to locate the <c>Kookerella.FsWordDsl</c>
    /// assembly - this has no opinion on that, since it depends on where the script ends up
    /// living (see <c>Document.generateScript</c>'s own F#-side doc comment).</summary>
    public static string GenerateFSharpScript(IEnumerable<string> referenceLines, string outputFileName, Document document) =>
        Fs.Document.generateScript(ListModule.OfSeq(referenceLines), outputFileName, DocumentConverter.ToFSharp(document));
}
