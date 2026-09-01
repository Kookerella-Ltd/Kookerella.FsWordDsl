namespace Kookerella.FsWordDsl.Interpreter

open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open Kookerella.FsWordDsl

/// The inverse of `ImageWriter` - a `Drawing` element (found inside a run) back into an
/// `ImageEntry`. Only the "inline, embedded (not linked), move-and-size-with-cell" shape
/// `ImageWriter` itself produces is recognized; anything else (free-floating anchors,
/// `<a:blip r:link>` rather than `r:embed`, formats beyond PNG/JPEG/GIF/BMP) is a documented
/// gap - see MAPPING.md, same "drop what isn't modeled" posture the rest of this DSL takes.
module ImageReader =

    let private formatOfContentType (contentType: string) : ImageFormat option =
        match contentType with
        | "image/png" -> Some Png
        | "image/jpeg" -> Some Jpeg
        | "image/gif" -> Some Gif
        | "image/bmp"
        | "image/x-ms-bmp" -> Some Bmp
        | _ -> None

    let tryReadImage (mainPart: MainDocumentPart) (drawing: Wordprocessing.Drawing) : ImageEntry option =
        drawing.Descendants<Drawing.Wordprocessing.Inline>()
        |> Seq.tryHead
        |> Option.bind (fun inl ->
            let blip = inl.Descendants<Drawing.Blip>() |> Seq.tryHead
            let extent = inl.Descendants<Drawing.Wordprocessing.Extent>() |> Seq.tryHead
            let docProps = inl.Descendants<Drawing.Wordprocessing.DocProperties>() |> Seq.tryHead

            match blip, extent with
            | Some blip, Some extent when not (isNull blip.Embed) ->
                let relId = blip.Embed.Value
                let part = mainPart.GetPartById(relId) :?> ImagePart

                match formatOfContentType part.ContentType with
                | None -> None
                | Some format ->
                    use stream = part.GetStream()
                    use mem = new MemoryStream()
                    stream.CopyTo(mem)

                    Some
                        { Data = mem.ToArray()
                          Format = format
                          WidthEmu = (if isNull extent.Cx then 0L else extent.Cx.Value)
                          HeightEmu = (if isNull extent.Cy then 0L else extent.Cy.Value)
                          AltText =
                            docProps
                            |> Option.bind (fun p -> p.Description |> Option.ofObj)
                            |> Option.map (fun v -> v.Value)
                            |> Option.filter (fun s -> s <> "") }
            | _ -> None)
