namespace Kookerella.FsWordDsl

open System.IO
open Kookerella.FsWordDsl.Interpreter

/// Friendly entry points: render a `Document` to a .docx/.docm file/stream (the
/// interpreter), and parse a .docx/.docm file/stream back into a `Document` (the reverse
/// transform).
module Document =

    let save (path: string) (doc: Document) : unit = Writer.saveToFile doc path

    let saveToStream (stream: Stream) (doc: Document) : unit = Writer.saveToStream doc stream

    let load (path: string) : Document = Reader.loadFromFile path

    let loadFromStream (stream: Stream) : Document = Reader.loadFromStream stream

    /// Renders `doc` as a self-contained F# script that, when run, rebuilds an equivalent
    /// file at `outputFileName`. `referenceLines` are whatever raw `#r` directives the
    /// caller needs so the script can locate the Kookerella.FsWordDsl assembly - this has no
    /// opinion on that, since it depends on where the script ends up living.
    let generateScript (referenceLines: string list) (outputFileName: string) (doc: Document) : string =
        CodeGen.generate referenceLines outputFileName doc
