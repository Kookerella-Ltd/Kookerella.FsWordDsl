namespace Kookerella.FsWordDsl.Interpreter

open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open Kookerella.FsWordDsl

/// `ImageEntry` -> the DrawingML an inline picture needs (`w:drawing`/`wp:inline`/`a:blip`) -
/// split out from `Writer.fs` the same way Excel splits its own `ImageWriter.fs`, even
/// though this is a smaller surface than Excel's cell-anchored version (no free-floating
/// anchor, no `DrawingsPart`/per-worksheet shared canvas to own - a Word inline image lives
/// entirely inside the one run that contains it). References the DrawingML/picture
/// namespaces via `open DocumentFormat.OpenXml` plus each nested namespace's own short name
/// (`Drawing.XXX`, `Drawing.Wordprocessing.XXX`, `Drawing.Pictures.XXX`) rather than a
/// `module` alias, since F# can't alias a namespace as a module - see `Writer.fs`'s own note.
module ImageWriter =

    let private imagePartType (format: ImageFormat) : PartTypeInfo =
        match format with
        | Png -> ImagePartType.Png
        | Jpeg -> ImagePartType.Jpeg
        | Gif -> ImagePartType.Gif
        | Bmp -> ImagePartType.Bmp

    /// `drawingId` only needs to be unique within the document (Word uses it purely as a
    /// display/debugging label, `DocProperties.Id`) - the caller threads a simple
    /// incrementing counter through, same idea as `Writer`'s bookmark/comment id counters.
    ///
    /// Every SINGLE-argument constructor call below that would otherwise wrap one child
    /// element (`Type(child)`) is deliberately written as `Type()` + `AppendChild` instead -
    /// F# always resolves a one-argument `OpenXmlElement`-typed constructor call to the
    /// SDK's `IEnumerable<OpenXmlElement>` overload rather than "one child to wrap" (every
    /// `OpenXmlCompositeElement` implements that interface over its own children), which
    /// silently produces an empty parent for a childless leaf argument or throws ("part of
    /// a tree") for a composite one that already has children of its own - see `Writer.fs`'s
    /// own note on the same gotcha. Multi-argument constructor calls (2+) are unaffected by
    /// this and are used normally throughout.
    let addImage (mainPart: MainDocumentPart) (drawingId: uint32) (img: ImageEntry) : Wordprocessing.Drawing =
        let imagePart = mainPart.AddImagePart(imagePartType img.Format)

        use stream = new MemoryStream(img.Data)
        imagePart.FeedData(stream)

        let relationshipId = mainPart.GetIdOfPart(imagePart)
        let name = sprintf "Picture %d" drawingId
        let description = img.AltText |> Option.defaultValue ""

        let preset = Drawing.PresetGeometry()
        preset.AppendChild(Drawing.AdjustValueList()) |> ignore
        preset.Preset <- EnumValue Drawing.ShapeTypeValues.Rectangle

        let stretch = Drawing.Stretch()
        stretch.AppendChild(Drawing.FillRectangle()) |> ignore

        let picture =
            Drawing.Pictures.Picture(
                Drawing.Pictures.NonVisualPictureProperties(
                    Drawing.Pictures.NonVisualDrawingProperties(Id = UInt32Value 0u, Name = StringValue name),
                    Drawing.Pictures.NonVisualPictureDrawingProperties()
                ),
                Drawing.Pictures.BlipFill(Drawing.Blip(Embed = StringValue relationshipId), stretch),
                Drawing.Pictures.ShapeProperties(
                    Drawing.Transform2D(Drawing.Offset(X = Int64Value 0L, Y = Int64Value 0L), Drawing.Extents(Cx = Int64Value img.WidthEmu, Cy = Int64Value img.HeightEmu)),
                    preset
                )
            )

        let graphicData = Drawing.GraphicData()
        graphicData.AppendChild(picture) |> ignore
        graphicData.Uri <- "http://schemas.openxmlformats.org/drawingml/2006/picture"

        let graphic = Drawing.Graphic()
        graphic.AppendChild(graphicData) |> ignore

        let nvGraphicFrameProps = Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties()
        nvGraphicFrameProps.AppendChild(Drawing.GraphicFrameLocks(NoChangeAspect = true)) |> ignore

        let inlineEl =
            Drawing.Wordprocessing.Inline(
                Drawing.Wordprocessing.Extent(Cx = Int64Value img.WidthEmu, Cy = Int64Value img.HeightEmu),
                Drawing.Wordprocessing.EffectExtent(LeftEdge = Int64Value 0L, TopEdge = Int64Value 0L, RightEdge = Int64Value 0L, BottomEdge = Int64Value 0L),
                Drawing.Wordprocessing.DocProperties(Id = UInt32Value drawingId, Name = StringValue name, Description = StringValue description),
                nvGraphicFrameProps,
                graphic
            )

        inlineEl.DistanceFromTop <- UInt32Value 0u
        inlineEl.DistanceFromBottom <- UInt32Value 0u
        inlineEl.DistanceFromLeft <- UInt32Value 0u
        inlineEl.DistanceFromRight <- UInt32Value 0u

        let drawing = Wordprocessing.Drawing()
        drawing.AppendChild(inlineEl) |> ignore
        drawing
