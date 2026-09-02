using Fs = Kookerella.FsWordDsl;

namespace Kookerella.CsWordDsl;

/// <summary>
/// The one place this wrapper does anything side-effecting - every other type in this
/// assembly is a pure, immutable value with no I/O methods of its own (see <see
/// cref="Document"/>'s own doc comment). Mirrors the F# core's own separation between the
/// <c>Document</c> data type and its <c>Document.save</c>/<c>load</c> module functions.
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
}
