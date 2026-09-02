namespace Kookerella.FsWordDsl.Interpreter

open System
open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open Kookerella.FsWordDsl
open Kookerella.FsWordDsl.Interpreter.StyleRegistry
open Kookerella.FsWordDsl.Interpreter.ImageWriter

/// DSL -> OOXML. The interpreter half of this DSL's round-trip pair (`Reader` is the
/// other) - never `open`s `DocumentFormat.OpenXml.Wordprocessing` directly (F# doesn't allow
/// aliasing a namespace as a module, so `open DocumentFormat.OpenXml` plus the
/// `Wordprocessing.XXX` qualified form is used throughout instead - the nested namespace's
/// own short name resolves once its parent is open) so the DSL's own natural type/case names
/// (`Paragraph`, `Table`, `Hyperlink`, `Bookmark`, `Comment`, ...) stay usable unqualified
/// with no collision - see `Model.fs`'s own note on this.
module Writer =

    let private pointsToTwips (pts: float) : int = int (Math.Round(pts * 20.0))
    let private pointsToTwipsU (pts: float) : uint32 = uint32 (Math.Round(pts * 20.0))

    /// Mutable state threaded through one `writeDocument` call - never shared across calls.
    /// `NextBookmarkId`/`NextCommentId`/`NextDrawingId` only need to be unique within this
    /// one document (they're OOXML's own internal ids, not anything the DSL exposes back to
    /// a caller), so a simple incrementing counter is enough, same idea as `ImageWriter`'s
    /// own doc comment on `drawingId`.
    type private Ctx =
        { MainPart: MainDocumentPart
          Comments: ResizeArray<Wordprocessing.Comment>
          /// `(id, note element)` pairs, in the order encountered - written to `word/
          /// footnotes.xml`/`endnotes.xml` (alongside the two boilerplate separator entries
          /// every real Word file carries) once the whole body has been walked.
          Footnotes: ResizeArray<int * Wordprocessing.Footnote>
          Endnotes: ResizeArray<int * Wordprocessing.Endnote>
          /// Name -> the numeric `w:id` assigned when its `BookmarkRangeStart` was written,
          /// removed once the matching `BookmarkRangeEnd` is written - lets the two ends of
          /// a cross-paragraph bookmark, written independently, still share OOXML's own
          /// required matching id even though only `BookmarkRangeStart` carries `name`.
          OpenBookmarkRanges: Collections.Generic.Dictionary<string, int>
          /// The comment-range equivalent of `OpenBookmarkRanges` - caller-chosen `id` ->
          /// the real OOXML comment id assigned when its `CommentRangeStart` was written.
          /// Comment ids are already `StringValue`s on the wire (see `Comment`'s own
          /// handling below), so this stores the assigned id as a string directly rather
          /// than needing the int<->string conversion `OpenBookmarkRanges` does.
          OpenCommentRanges: Collections.Generic.Dictionary<string, string>
          mutable NextBookmarkId: int
          mutable NextCommentId: int
          mutable NextDrawingId: uint32
          mutable NextFootnoteId: int
          mutable NextEndnoteId: int
          mutable NextRevisionId: int
          /// Whether the `Inline` currently being written sits inside a `TrackedChange`
          /// with `Kind = Deleted` - a run's own text uses `w:delText` instead of `w:t`
          /// there (same run properties/styling either way). Save-and-restore around each
          /// `Deleted` case's own recursion (not a hard reset to `false`), so a `Deleted`
          /// nested inside another `Deleted` - unusual, but structurally possible - doesn't
          /// incorrectly flip back to "not deleted" partway through.
          mutable InsideDeletion: bool
          /// Whether ANY section uses an `Even` header/footer - `<w:evenAndOddHeaders/>` is
          /// a document-wide `settings.xml` flag, not a per-section `sectPr` child, unlike
          /// `<w:titlePg/>` (which genuinely is per-section).
          mutable NeedsEvenAndOddHeaders: bool }

    // --- Numbering ------------------------------------------------------------------------

    let private numberFormatKindToW (k: NumberFormatKind) : Wordprocessing.NumberFormatValues =
        match k with
        | BulletFormat _ -> Wordprocessing.NumberFormatValues.Bullet
        | DecimalFormat -> Wordprocessing.NumberFormatValues.Decimal
        | LowerLetterFormat -> Wordprocessing.NumberFormatValues.LowerLetter
        | UpperLetterFormat -> Wordprocessing.NumberFormatValues.UpperLetter
        | LowerRomanFormat -> Wordprocessing.NumberFormatValues.LowerRoman
        | UpperRomanFormat -> Wordprocessing.NumberFormatValues.UpperRoman
        | OtherFormat raw -> Wordprocessing.NumberFormatValues raw

    let private levelToW (ilvl: int) (level: ListLevel) : Wordprocessing.Level =
        let lvl = Wordprocessing.Level(LevelIndex = Int32Value ilvl)
        lvl.NumberingFormat <- Wordprocessing.NumberingFormat(Val = EnumValue(numberFormatKindToW level.Format))
        lvl.LevelText <- Wordprocessing.LevelText(Val = StringValue level.Text)
        lvl.LevelJustification <- Wordprocessing.LevelJustification(Val = EnumValue Wordprocessing.LevelJustificationValues.Left)
        level.StartAt |> Option.iter (fun s -> lvl.StartNumberingValue <- Wordprocessing.StartNumberingValue(Val = Int32Value s))

        if level.IndentLeft.IsSome || level.HangingIndent.IsSome then
            let ind = Wordprocessing.Indentation()
            level.IndentLeft |> Option.iter (fun v -> ind.Left <- StringValue(string (pointsToTwips v)))
            level.HangingIndent |> Option.iter (fun v -> ind.Hanging <- StringValue(string (pointsToTwips v)))
            lvl.PreviousParagraphProperties <- Wordprocessing.PreviousParagraphProperties(Indentation = ind)

        match level.Format with
        | BulletFormat(_, font) ->
            lvl.NumberingSymbolRunProperties <- Wordprocessing.NumberingSymbolRunProperties(RunFonts = Wordprocessing.RunFonts(Ascii = StringValue font, HighAnsi = StringValue font))
        | _ -> ()

        lvl

    /// `AbstractNum`s must all precede `NumberingInstance`s in `numbering.xml` - reuses each
    /// definition's own `Id` for both the abstract numbering id and the instance's `numId`
    /// (this DSL collapses Word's own abstract/instance indirection away, see `Numbering.fs`'s
    /// own doc comment).
    let private numberingToOpenXml (definitions: NumberingDefinition list) : Wordprocessing.Numbering =
        let abstractNums =
            definitions
            |> List.map (fun d ->
                let absNum = Wordprocessing.AbstractNum(AbstractNumberId = Int32Value d.Id)
                d.Levels |> List.iteri (fun i lvl -> absNum.AppendChild(levelToW i lvl) |> ignore)
                absNum :> OpenXmlElement)

        let instances =
            definitions
            |> List.map (fun d ->
                let inst = Wordprocessing.NumberingInstance(NumberID = Int32Value d.Id)
                inst.AppendChild(Wordprocessing.AbstractNumId(Val = Int32Value d.Id)) |> ignore
                inst :> OpenXmlElement)

        Wordprocessing.Numbering(abstractNums @ instances)

    // --- Borders --------------------------------------------------------------------------
    //
    // `borderSideToTop`/`ToBottom`/`ToLeft`/`ToRight`/`ToInsideH`/`ToInsideV` live in
    // `StyleRegistry.fs` now - shared with paragraph borders (`w:pBdr`), which reuse the
    // very same SDK element classes.

    let private tableBordersToW (b: TableBorders) : Wordprocessing.TableBorders =
        let tb = Wordprocessing.TableBorders()
        b.Outer.Top |> Option.iter (fun s -> tb.TopBorder <- borderSideToTop s)
        b.Outer.Bottom |> Option.iter (fun s -> tb.BottomBorder <- borderSideToBottom s)
        b.Outer.Left |> Option.iter (fun s -> tb.LeftBorder <- borderSideToLeft s)
        b.Outer.Right |> Option.iter (fun s -> tb.RightBorder <- borderSideToRight s)
        b.InsideHorizontal |> Option.iter (fun s -> tb.InsideHorizontalBorder <- borderSideToInsideH s)
        b.InsideVertical |> Option.iter (fun s -> tb.InsideVerticalBorder <- borderSideToInsideV s)
        tb

    let private tableCellBordersToW (b: TableBorders) : Wordprocessing.TableCellBorders =
        let tcb = Wordprocessing.TableCellBorders()
        b.Outer.Top |> Option.iter (fun s -> tcb.TopBorder <- borderSideToTop s)
        b.Outer.Bottom |> Option.iter (fun s -> tcb.BottomBorder <- borderSideToBottom s)
        b.Outer.Left |> Option.iter (fun s -> tcb.LeftBorder <- borderSideToLeft s)
        b.Outer.Right |> Option.iter (fun s -> tcb.RightBorder <- borderSideToRight s)
        tcb

    // --- Inline content ---------------------------------------------------------------------

    let private textRun (ctx: Ctx) (text: string) (style: RunStyle option) (styleId: string option) : Wordprocessing.Run =
        let r = Wordprocessing.Run()
        runPropertiesOf style styleId |> Option.iter (fun p -> r.RunProperties <- p)
        let preserveSpace = text.StartsWith(" ") || text.EndsWith(" ") || text.Contains("\t")

        // A run inside `w:del` (`ctx.InsideDeletion`) uses `w:delText` instead of `w:t` for
        // its own text - schema-required, not just convention (`OpenXmlValidator` flags a
        // plain `w:t` there). Same run properties/styling either way.
        if ctx.InsideDeletion then
            let t = Wordprocessing.DeletedText(text)
            if preserveSpace then t.Space <- EnumValue SpaceProcessingModeValues.Preserve
            r.AppendChild(t) |> ignore
        else
            let t = Wordprocessing.Text(text)
            if preserveSpace then t.Space <- EnumValue SpaceProcessingModeValues.Preserve
            r.AppendChild(t) |> ignore

        r

    /// A `Wordprocessing.Run` wrapping exactly one child. NOT `Wordprocessing.Run(child)`
    /// (a single-argument constructor call) - F# always resolves that to the SDK's
    /// `IEnumerable<OpenXmlElement>` constructor overload rather than "one child to wrap"
    /// (every `OpenXmlCompositeElement`, including leaf-ish ones like `Break`/`TabChar`,
    /// implements that interface over its own children), which silently produces an EMPTY
    /// run for a childless leaf element, or throws ("part of a tree") for a composite one
    /// that already has children - see this module's own note by `Document`'s construction
    /// below for the same gotcha at the top level. `AppendChild` sidesteps it entirely.
    let private runWith (child: OpenXmlElement) : Wordprocessing.Run =
        let r = Wordprocessing.Run()
        r.AppendChild(child) |> ignore
        r

    /// Prepends a note-reference-mark run (`w:footnoteRef`/`w:endnoteRef`, styled via
    /// `markerStyleId`) to the FIRST paragraph of a note body's already-built OOXML
    /// elements - inserted right after that paragraph's own `w:pPr` if it has one (which
    /// must stay the first child per schema), otherwise as the new first child outright. A
    /// note body with no paragraph at all (an empty `content` list, or one starting with a
    /// table) gets a synthetic leading paragraph to carry the marker, matching what real
    /// Word itself does.
    let private insertNoteMarker (markerStyleId: string) (mark: OpenXmlElement) (elements: OpenXmlElement list) : OpenXmlElement list =
        let markerRun = Wordprocessing.Run()
        markerRun.AppendChild(Wordprocessing.RunProperties(RunStyle = Wordprocessing.RunStyle(Val = StringValue markerStyleId))) |> ignore
        markerRun.AppendChild(mark) |> ignore

        match elements with
        | (:? Wordprocessing.Paragraph as firstPara) :: rest ->
            (match firstPara.ParagraphProperties with
             | null -> firstPara.PrependChild(markerRun) |> ignore
             | pPr -> firstPara.InsertAfter(markerRun, pPr) |> ignore)

            (firstPara :> OpenXmlElement) :: rest
        | _ ->
            let p = Wordprocessing.Paragraph()
            p.AppendChild(markerRun) |> ignore
            (p :> OpenXmlElement) :: elements

    /// Builds the boilerplate separator entries (`id = -1` "separator", `id = 0`
    /// "continuationSeparator") every real Word-authored `footnotes.xml`/`endnotes.xml`
    /// carries, for the horizontal separator line above the notes area - schema-optional,
    /// but real Word relies on their presence for correct rendering, so `Writer` always
    /// includes them whenever it writes either part at all.
    let private separatorParagraph (mark: OpenXmlElement) : Wordprocessing.Paragraph =
        let p = Wordprocessing.Paragraph()
        let r = Wordprocessing.Run()
        r.AppendChild(mark) |> ignore
        p.AppendChild(r) |> ignore
        p

    let private paragraphPropertiesFull
        (ctx: Ctx)
        (styleId: string option)
        (format: ParagraphFormat option)
        (numbering: (int * int) option)
        (markRevision: Revision option)
        : Wordprocessing.ParagraphProperties option =
        let basePr = paragraphPropertiesOf styleId format

        let withNumbering =
            match numbering with
            | None -> basePr
            | Some(numId, level) ->
                let pPr = basePr |> Option.defaultValue (Wordprocessing.ParagraphProperties())
                pPr.NumberingProperties <- Wordprocessing.NumberingProperties(Wordprocessing.NumberingLevelReference(Val = Int32Value level), Wordprocessing.NumberingId(Val = Int32Value numId))
                Some pPr

        // The paragraph MARK's own revision (`w:pPr/w:rPr/w:ins`|`w:del`) - distinct from
        // any `TrackedChange` wrapping the paragraph's own `Inlines`, which marks the
        // *content* rather than the mark. `Inserted`/`Deleted` here are the SDK's leaf flag
        // classes (author/date/id attributes only) - a different pair of classes from
        // `InsertedRun`/`DeletedRun`, which wrap actual run content elsewhere in this file.
        match markRevision with
        | None -> withNumbering
        | Some revision ->
            let pPr = withNumbering |> Option.defaultValue (Wordprocessing.ParagraphProperties())
            let id = ctx.NextRevisionId
            ctx.NextRevisionId <- id + 1
            let rPr = Wordprocessing.ParagraphMarkRunProperties()
            let dateVal = DateTimeValue(defaultArg revision.Date DateTime.Now)

            match revision.Kind with
            | Inserted ->
                let flag = Wordprocessing.Inserted(Id = StringValue(string id), Author = StringValue revision.Author)
                flag.Date <- dateVal
                rPr.Inserted <- flag
            | Deleted ->
                let flag = Wordprocessing.Deleted(Id = StringValue(string id), Author = StringValue revision.Author)
                flag.Date <- dateVal
                rPr.Deleted <- flag

            pPr.ParagraphMarkRunProperties <- rPr
            Some pPr

    /// Assigns a fresh comment id, builds its `word/comments.xml` metadata entry, and adds
    /// it to `ctx.Comments` - shared by the wrapping `Comment` case and `CommentRangeStart`
    /// below, which differ only in how the surrounding range/content is written. Returns
    /// the assigned id (as a string, matching `w:id`'s own wire type).
    let private addCommentMetadata (ctx: Ctx) (author: string) (initials: string option) (date: DateTime option) (text: string) : string =
        let id = ctx.NextCommentId
        ctx.NextCommentId <- id + 1
        let idStr = string id

        let cmt = Wordprocessing.Comment(Id = StringValue idStr, Author = StringValue author)
        cmt.Date <- DateTimeValue(defaultArg date DateTime.Now)
        initials |> Option.iter (fun i -> cmt.Initials <- StringValue i)
        let commentPara = Wordprocessing.Paragraph()
        commentPara.AppendChild(runWith (Wordprocessing.Text(text))) |> ignore
        cmt.AppendChild(commentPara) |> ignore
        ctx.Comments.Add(cmt)

        idStr

    /// Builds a `w:sdtPr` for one `ContentControlProps` - `SdtProperties` itself exposes no
    /// named child properties in the SDK (confirmed by reflection, unlike almost every
    /// other composite element this DSL constructs), so its children are appended
    /// positionally like `w:tblGrid`'s columns rather than assigned. `CheckBoxControl` is
    /// the one case that reaches outside `Wordprocessing` - `Office2010.Word.
    /// SdtContentCheckBox`/`.Checked` on the wire (`w14:checkbox`/`w14:checked`), a
    /// DIFFERENT `OnOffValues` enum from `Wordprocessing.OnOffValues` (confirmed by
    /// reflection - `Checked.Val : EnumValue<Office2010.Word.OnOffValues>`), so it must be
    /// qualified explicitly rather than via the ambient `Wordprocessing.OnOffValues` this
    /// file uses everywhere else.
    let private lockToW (lock: ContentControlLock) : Wordprocessing.LockingValues =
        match lock with
        | LockDeletion -> Wordprocessing.LockingValues.SdtLocked
        | LockContentEditing -> Wordprocessing.LockingValues.ContentLocked
        | LockDeletionAndContentEditing -> Wordprocessing.LockingValues.SdtContentLocked

    let private contentControlPropsToW (props: ContentControlProps) : Wordprocessing.SdtProperties =
        let sdtPr = Wordprocessing.SdtProperties()
        props.Alias |> Option.iter (fun a -> sdtPr.AppendChild(Wordprocessing.SdtAlias(Val = StringValue a)) |> ignore)
        props.Tag |> Option.iter (fun t -> sdtPr.AppendChild(Wordprocessing.Tag(Val = StringValue t)) |> ignore)
        props.Lock |> Option.iter (fun l -> sdtPr.AppendChild(Wordprocessing.Lock(Val = EnumValue(lockToW l))) |> ignore)

        match props.Type with
        | RichTextControl -> ()
        | PlainTextControl multiLine ->
            let t = Wordprocessing.SdtContentText()
            if multiLine then t.MultiLine <- OnOffValue true
            sdtPr.AppendChild(t) |> ignore
        | DropDownControl(items, editable) ->
            let listItems = items |> List.map (fun (display, value) -> Wordprocessing.ListItem(DisplayText = StringValue display, Value = StringValue value) :> OpenXmlElement)

            if editable then
                sdtPr.AppendChild(Wordprocessing.SdtContentComboBox(listItems)) |> ignore
            else
                sdtPr.AppendChild(Wordprocessing.SdtContentDropDownList(listItems)) |> ignore
        | DateControl(fullDate, format) ->
            let d = Wordprocessing.SdtContentDate()
            fullDate |> Option.iter (fun dt -> d.FullDate <- DateTimeValue dt)
            format |> Option.iter (fun f -> d.DateFormat <- Wordprocessing.DateFormat(Val = StringValue f))
            sdtPr.AppendChild(d) |> ignore
        | CheckBoxControl(checked_, checkedSymbol, uncheckedSymbol) ->
            let cb = Office2010.Word.SdtContentCheckBox()
            let onOff = if checked_ then Office2010.Word.OnOffValues.One else Office2010.Word.OnOffValues.Zero
            cb.Checked <- Office2010.Word.Checked(Val = EnumValue onOff)
            checkedSymbol |> Option.iter (fun (font, code) -> cb.CheckedState <- Office2010.Word.CheckedState(Font = StringValue font, Val = HexBinaryValue code))
            uncheckedSymbol |> Option.iter (fun (font, code) -> cb.UncheckedState <- Office2010.Word.UncheckedState(Font = StringValue font, Val = HexBinaryValue code))
            sdtPr.AppendChild(cb) |> ignore

        sdtPr

    let rec private inlineToElements (ctx: Ctx) (inl: Inline) : OpenXmlElement list =
        match inl with
        | Run(text, style, styleId) -> [ textRun ctx text style styleId :> OpenXmlElement ]
        | LineBreak -> [ runWith (Wordprocessing.Break(Type = EnumValue Wordprocessing.BreakValues.TextWrapping)) :> OpenXmlElement ]
        | Tab -> [ runWith (Wordprocessing.TabChar()) :> OpenXmlElement ]
        | PageBreak -> [ runWith (Wordprocessing.Break(Type = EnumValue Wordprocessing.BreakValues.Page)) :> OpenXmlElement ]
        | Image img ->
            let id = ctx.NextDrawingId
            ctx.NextDrawingId <- id + 1u
            [ runWith (addImage ctx.MainPart id img) :> OpenXmlElement ]
        | Hyperlink(target, runs, tooltip) ->
            let children = runs |> List.collect (inlineToElements ctx)

            let hl =
                match target with
                | ExternalUrl url ->
                    let rel = ctx.MainPart.AddHyperlinkRelationship(Uri(url, UriKind.RelativeOrAbsolute), true)
                    Wordprocessing.Hyperlink(Id = StringValue rel.Id, History = OnOffValue true)
                | InternalBookmark name -> Wordprocessing.Hyperlink(Anchor = StringValue name, History = OnOffValue true)

            tooltip |> Option.iter (fun tt -> hl.Tooltip <- StringValue tt)
            children |> List.iter (hl.AppendChild >> ignore)
            [ hl :> OpenXmlElement ]
        | Bookmark(name, content) ->
            let id = ctx.NextBookmarkId
            ctx.NextBookmarkId <- id + 1
            let startEl = Wordprocessing.BookmarkStart(Id = StringValue(string id), Name = StringValue name) :> OpenXmlElement
            let endEl = Wordprocessing.BookmarkEnd(Id = StringValue(string id)) :> OpenXmlElement
            let contentEls = content |> List.collect (inlineToElements ctx)
            [ startEl ] @ contentEls @ [ endEl ]
        | BookmarkRangeStart name ->
            let id = ctx.NextBookmarkId
            ctx.NextBookmarkId <- id + 1
            ctx.OpenBookmarkRanges.[name] <- id
            [ Wordprocessing.BookmarkStart(Id = StringValue(string id), Name = StringValue name) :> OpenXmlElement ]
        | BookmarkRangeEnd name ->
            match ctx.OpenBookmarkRanges.TryGetValue name with
            | true, id ->
                ctx.OpenBookmarkRanges.Remove(name) |> ignore
                [ Wordprocessing.BookmarkEnd(Id = StringValue(string id)) :> OpenXmlElement ]
            | false, _ -> failwithf "BookmarkRangeEnd %s has no matching BookmarkRangeStart" name
        | Comment(author, initials, date, text, content) ->
            let idStr = addCommentMetadata ctx author initials date text
            let startEl = Wordprocessing.CommentRangeStart(Id = StringValue idStr) :> OpenXmlElement
            let endEl = Wordprocessing.CommentRangeEnd(Id = StringValue idStr) :> OpenXmlElement
            let refRun = runWith (Wordprocessing.CommentReference(Id = StringValue idStr)) :> OpenXmlElement
            let contentEls = content |> List.collect (inlineToElements ctx)
            [ startEl ] @ contentEls @ [ endEl; refRun ]
        | CommentRangeStart(callerId, author, initials, date, text) ->
            let idStr = addCommentMetadata ctx author initials date text
            ctx.OpenCommentRanges.[callerId] <- idStr
            [ Wordprocessing.CommentRangeStart(Id = StringValue idStr) :> OpenXmlElement ]
        | CommentRangeEnd callerId ->
            match ctx.OpenCommentRanges.TryGetValue callerId with
            | true, idStr ->
                ctx.OpenCommentRanges.Remove(callerId) |> ignore
                [ Wordprocessing.CommentRangeEnd(Id = StringValue idStr) :> OpenXmlElement
                  runWith (Wordprocessing.CommentReference(Id = StringValue idStr)) :> OpenXmlElement ]
            | false, _ -> failwithf "CommentRangeEnd %s has no matching CommentRangeStart" callerId
        | TrackedChange(revision, content) ->
            let id = ctx.NextRevisionId
            ctx.NextRevisionId <- id + 1
            let dateVal = DateTimeValue(defaultArg revision.Date DateTime.Now)

            match revision.Kind with
            | Inserted ->
                let wrapper = Wordprocessing.InsertedRun(Id = StringValue(string id), Author = StringValue revision.Author)
                wrapper.Date <- dateVal
                let contentEls = content |> List.collect (inlineToElements ctx)
                contentEls |> List.iter (wrapper.AppendChild >> ignore)
                [ wrapper :> OpenXmlElement ]
            | Deleted ->
                let wrapper = Wordprocessing.DeletedRun(Id = StringValue(string id), Author = StringValue revision.Author)
                wrapper.Date <- dateVal
                let wasInsideDeletion = ctx.InsideDeletion
                ctx.InsideDeletion <- true
                let contentEls = content |> List.collect (inlineToElements ctx)
                ctx.InsideDeletion <- wasInsideDeletion
                contentEls |> List.iter (wrapper.AppendChild >> ignore)
                [ wrapper :> OpenXmlElement ]
        | InlineContentControl(props, content) ->
            let sdt = Wordprocessing.SdtRun()
            sdt.SdtProperties <- contentControlPropsToW props
            let sdtContent = Wordprocessing.SdtContentRun()
            content |> List.collect (inlineToElements ctx) |> List.iter (sdtContent.AppendChild >> ignore)
            sdt.SdtContentRun <- sdtContent
            [ sdt :> OpenXmlElement ]
        | Field(instruction, cachedResult) ->
            let sf = Wordprocessing.SimpleField(Instruction = StringValue instruction)
            cachedResult |> Option.iter (fun c -> sf.AppendChild(runWith (Wordprocessing.Text(c))) |> ignore)
            [ sf :> OpenXmlElement ]
        | Footnote content ->
            let id = ctx.NextFootnoteId
            ctx.NextFootnoteId <- id + 1

            let note = Wordprocessing.Footnote(Id = IntegerValue id)

            content
            |> List.map (blockToW ctx)
            |> insertNoteMarker "FootnoteReference" (Wordprocessing.FootnoteReferenceMark())
            |> List.iter (fun el -> note.AppendChild(el) |> ignore)

            ctx.Footnotes.Add(id, note)

            let refRun = Wordprocessing.Run()
            refRun.AppendChild(Wordprocessing.RunProperties(RunStyle = Wordprocessing.RunStyle(Val = StringValue "FootnoteReference"))) |> ignore
            refRun.AppendChild(Wordprocessing.FootnoteReference(Id = IntegerValue id)) |> ignore
            [ refRun :> OpenXmlElement ]
        | Endnote content ->
            let id = ctx.NextEndnoteId
            ctx.NextEndnoteId <- id + 1

            let note = Wordprocessing.Endnote(Id = IntegerValue id)

            content
            |> List.map (blockToW ctx)
            |> insertNoteMarker "EndnoteReference" (Wordprocessing.EndnoteReferenceMark())
            |> List.iter (fun el -> note.AppendChild(el) |> ignore)

            ctx.Endnotes.Add(id, note)

            let refRun = Wordprocessing.Run()
            refRun.AppendChild(Wordprocessing.RunProperties(RunStyle = Wordprocessing.RunStyle(Val = StringValue "EndnoteReference"))) |> ignore
            refRun.AppendChild(Wordprocessing.EndnoteReference(Id = IntegerValue id)) |> ignore
            [ refRun :> OpenXmlElement ]

    // --- Paragraphs / tables ----------------------------------------------------------------
    //
    // Continues the same `rec ... and ...` chain `inlineToElements` above started, rather
    // than a separate group - `Footnote`/`Endnote` above call `blockToW` (a note's body is a
    // `Block list`), which calls `paragraphToW`, which calls `inlineToElements` for its own
    // `Inlines`, closing the cycle the model-level `Inline`/`Block` mutual recursion implies.

    and private paragraphToW (ctx: Ctx) (p: Paragraph) : Wordprocessing.Paragraph =
        let para = Wordprocessing.Paragraph()
        paragraphPropertiesFull ctx p.StyleId p.Format p.Numbering p.MarkRevision |> Option.iter (fun pPr -> para.ParagraphProperties <- pPr)
        p.Inlines |> List.collect (inlineToElements ctx) |> List.iter (para.AppendChild >> ignore)
        para

    and private blockToW (ctx: Ctx) (block: Block) : OpenXmlElement =
        match block with
        | ParagraphBlock p -> paragraphToW ctx p :> OpenXmlElement
        | TableBlock t -> tableEntryToW ctx t :> OpenXmlElement
        | ContentControlBlock(props, content) ->
            let sdt = Wordprocessing.SdtBlock()
            sdt.SdtProperties <- contentControlPropsToW props
            let sdtContent = Wordprocessing.SdtContentBlock()
            content |> List.iter (fun b -> sdtContent.AppendChild(blockToW ctx b) |> ignore)
            sdt.SdtContentBlock <- sdtContent
            sdt :> OpenXmlElement

    and private tableCellToW (ctx: Ctx) (colWidthTwips: int) (cell: TableCell) : Wordprocessing.TableCell =
        let tc = Wordprocessing.TableCell()
        let tcPr = Wordprocessing.TableCellProperties()

        let widthTwips =
            cell.Props.Width |> Option.map pointsToTwips |> Option.defaultValue colWidthTwips

        tcPr.TableCellWidth <- Wordprocessing.TableCellWidth(Type = EnumValue Wordprocessing.TableWidthUnitValues.Dxa, Width = StringValue(string widthTwips))
        cell.Props.GridSpan |> Option.iter (fun n -> tcPr.GridSpan <- Wordprocessing.GridSpan(Val = Int32Value n))

        cell.Props.VerticalMerge
        |> Option.iter (fun vm ->
            let v =
                match vm with
                | RestartMerge -> Wordprocessing.MergedCellValues.Restart
                | ContinueMerge -> Wordprocessing.MergedCellValues.Continue

            tcPr.VerticalMerge <- Wordprocessing.VerticalMerge(Val = EnumValue v))

        cell.Props.Shading
        |> Option.iter (fun c ->
            let sh = Wordprocessing.Shading(Val = EnumValue Wordprocessing.ShadingPatternValues.Clear, Color = StringValue "auto", Fill = StringValue(colorToHex c))
            applyThemeToShadingFill sh c
            tcPr.Shading <- sh)

        cell.Props.Borders |> Option.iter (fun b -> tcPr.TableCellBorders <- tableCellBordersToW b)
        cell.Props.Margins |> Option.iter (fun m -> tcPr.TableCellMargin <- cellMarginsToTcMar m)
        tc.TableCellProperties <- tcPr

        if cell.Content.IsEmpty then
            tc.AppendChild(Wordprocessing.Paragraph()) |> ignore
        else
            cell.Content |> List.iter (fun b -> tc.AppendChild(blockToW ctx b) |> ignore)

        tc

    and private cellMarginsToW (m: CellMargins) : Wordprocessing.TableCellMarginDefault =
        let tcMar = Wordprocessing.TableCellMarginDefault()
        m.Top |> Option.iter (fun v -> tcMar.TopMargin <- Wordprocessing.TopMargin(Width = StringValue(string (pointsToTwips v)), Type = EnumValue Wordprocessing.TableWidthUnitValues.Dxa))
        m.Bottom |> Option.iter (fun v -> tcMar.BottomMargin <- Wordprocessing.BottomMargin(Width = StringValue(string (pointsToTwips v)), Type = EnumValue Wordprocessing.TableWidthUnitValues.Dxa))
        m.Left |> Option.iter (fun v -> tcMar.TableCellLeftMargin <- Wordprocessing.TableCellLeftMargin(Width = Int16Value(int16 (pointsToTwips v)), Type = EnumValue Wordprocessing.TableWidthValues.Dxa))
        m.Right |> Option.iter (fun v -> tcMar.TableCellRightMargin <- Wordprocessing.TableCellRightMargin(Width = Int16Value(int16 (pointsToTwips v)), Type = EnumValue Wordprocessing.TableWidthValues.Dxa))
        tcMar

    /// The per-cell equivalent of `cellMarginsToW` (`w:tcPr/w:tcMar`, overriding the
    /// table's own default) - a different, if similarly-shaped, set of SDK child element
    /// classes than the table-wide default uses (`LeftMargin`/`RightMargin` here, not
    /// `TableCellLeftMargin`/`TableCellRightMargin`), confirmed by reflection same as
    /// everywhere else this DSL constructs OOXML elements.
    and private cellMarginsToTcMar (m: CellMargins) : Wordprocessing.TableCellMargin =
        let tcMar = Wordprocessing.TableCellMargin()
        m.Top |> Option.iter (fun v -> tcMar.TopMargin <- Wordprocessing.TopMargin(Width = StringValue(string (pointsToTwips v)), Type = EnumValue Wordprocessing.TableWidthUnitValues.Dxa))
        m.Bottom |> Option.iter (fun v -> tcMar.BottomMargin <- Wordprocessing.BottomMargin(Width = StringValue(string (pointsToTwips v)), Type = EnumValue Wordprocessing.TableWidthUnitValues.Dxa))
        m.Left |> Option.iter (fun v -> tcMar.LeftMargin <- Wordprocessing.LeftMargin(Width = StringValue(string (pointsToTwips v)), Type = EnumValue Wordprocessing.TableWidthUnitValues.Dxa))
        m.Right |> Option.iter (fun v -> tcMar.RightMargin <- Wordprocessing.RightMargin(Width = StringValue(string (pointsToTwips v)), Type = EnumValue Wordprocessing.TableWidthUnitValues.Dxa))
        tcMar

    and private tableEntryToW (ctx: Ctx) (t: TableEntry) : Wordprocessing.Table =
        let table = Wordprocessing.Table()
        let tblPr = Wordprocessing.TableProperties()

        t.Style
        |> Option.iter (fun s ->
            tblPr.TableStyle <- Wordprocessing.TableStyle(Val = StringValue s.Name)

            let look = Wordprocessing.TableLook(Val = HexBinaryValue "0000")
            look.FirstRow <- OnOffValue s.FirstRowBanding
            look.LastRow <- OnOffValue s.LastRowBanding
            look.NoHorizontalBand <- OnOffValue(not s.BandedRows)
            look.NoVerticalBand <- OnOffValue(not s.BandedColumns)
            tblPr.TableLook <- look)

        t.Borders |> Option.iter (fun b -> tblPr.TableBorders <- tableBordersToW b)
        t.CellMargins |> Option.iter (fun m -> tblPr.TableCellMarginDefault <- cellMarginsToW m)
        table.AppendChild(tblPr) |> ignore

        let widthsTwips = t.ColumnWidths |> List.map pointsToTwips
        let grid = Wordprocessing.TableGrid(widthsTwips |> List.map (fun w -> Wordprocessing.GridColumn(Width = StringValue(string w)) :> OpenXmlElement))
        table.AppendChild(grid) |> ignore

        t.Rows
        |> List.iter (fun row ->
            let tr = Wordprocessing.TableRow()

            if row.Height.IsSome || row.RepeatAsHeader then
                let trPr = Wordprocessing.TableRowProperties()
                row.Height |> Option.iter (fun h -> trPr.AppendChild(Wordprocessing.TableRowHeight(Val = UInt32Value(pointsToTwipsU h))) |> ignore)

                if row.RepeatAsHeader then
                    trPr.AppendChild(Wordprocessing.TableHeader()) |> ignore

                tr.AppendChild(trPr) |> ignore

            row.Cells
            |> List.iteri (fun i cell ->
                let colWidth = if i < widthsTwips.Length then widthsTwips.[i] else 0
                tr.AppendChild(tableCellToW ctx colWidth cell) |> ignore)

            table.AppendChild(tr) |> ignore)

        table

    // --- Page setup / headers & footers -----------------------------------------------------

    let private namedPageSizeTwipsPortrait (size: PageSize) : int * int =
        match size with
        | Letter -> 12240, 15840
        | Legal -> 12240, 20160
        | A4 -> 11906, 16838
        | A3 -> 16838, 23811
        | OtherPageSize _ -> 12240, 15840
        | CustomPageSize(w, h) -> pointsToTwips w, pointsToTwips h

    let private pageSizeToW (size: PageSize) (orientation: PageOrientation) : Wordprocessing.PageSize =
        let w, h = namedPageSizeTwipsPortrait size
        let w, h = if orientation = Landscape then h, w else w, h
        let ps = Wordprocessing.PageSize(Width = UInt32Value(uint32 w), Height = UInt32Value(uint32 h))

        if orientation = Landscape then
            ps.Orient <- EnumValue Wordprocessing.PageOrientationValues.Landscape

        match size with
        | OtherPageSize code -> ps.Code <- UInt16Value(uint16 code)
        | _ -> ()

        ps

    let private pageMarginsToW (m: PageMargins) : Wordprocessing.PageMargin =
        Wordprocessing.PageMargin(
            Top = Int32Value(pointsToTwips m.Top),
            Bottom = Int32Value(pointsToTwips m.Bottom),
            Left = UInt32Value(pointsToTwipsU m.Left),
            Right = UInt32Value(pointsToTwipsU m.Right),
            Header = UInt32Value(pointsToTwipsU m.Header),
            Footer = UInt32Value(pointsToTwipsU m.Footer),
            Gutter = UInt32Value(pointsToTwipsU m.Gutter)
        )

    /// Adds a `HeaderPart`/`FooterPart` for one `Block list`, returning the relationship id
    /// a `HeaderReference`/`FooterReference` needs.
    let private addHeaderPart (ctx: Ctx) (blocks: Block list) : string =
        let part = ctx.MainPart.AddNewPart<HeaderPart>()
        part.Header <- Wordprocessing.Header(blocks |> List.map (blockToW ctx))
        ctx.MainPart.GetIdOfPart(part)

    let private addFooterPart (ctx: Ctx) (blocks: Block list) : string =
        let part = ctx.MainPart.AddNewPart<FooterPart>()
        part.Footer <- Wordprocessing.Footer(blocks |> List.map (blockToW ctx))
        ctx.MainPart.GetIdOfPart(part)

    let private sectionBreakTypeToW (t: SectionBreakType) : Wordprocessing.SectionMarkValues option =
        match t with
        // Not written at all - "next page" is Word's own default when `<w:type>` is
        // absent, same "only write what differs from the default" posture the rest of
        // this DSL takes (e.g. `BorderSide.Width = None` uses OOXML's own default weight).
        | NextPageBreak -> None
        | ContinuousBreak -> Some Wordprocessing.SectionMarkValues.Continuous
        | EvenPageBreak -> Some Wordprocessing.SectionMarkValues.EvenPage
        | OddPageBreak -> Some Wordprocessing.SectionMarkValues.OddPage

    /// Builds the `<w:sectPr>` for one section, wiring header/footer relationship ids and
    /// the auto-flag(s) they each require (see `Model.HeaderFooterSet`'s own doc comment).
    let private noteNumberRestartToW (r: NoteNumberRestart) : Wordprocessing.RestartNumberValues =
        match r with
        | ContinuousRestart -> Wordprocessing.RestartNumberValues.Continuous
        | RestartEachSection -> Wordprocessing.RestartNumberValues.EachSection
        | RestartEachPage -> Wordprocessing.RestartNumberValues.EachPage

    let private sectionPropertiesToW (ctx: Ctx) (props: SectionProperties) : Wordprocessing.SectionProperties =
        let sectPr = Wordprocessing.SectionProperties()
        let mutable titlePage = false

        props.Header
        |> Option.iter (fun h ->
            h.First |> Option.iter (fun blocks -> sectPr.AppendChild(Wordprocessing.HeaderReference(Type = EnumValue Wordprocessing.HeaderFooterValues.First, Id = StringValue(addHeaderPart ctx blocks))) |> ignore; titlePage <- true)
            h.Default |> Option.iter (fun blocks -> sectPr.AppendChild(Wordprocessing.HeaderReference(Type = EnumValue Wordprocessing.HeaderFooterValues.Default, Id = StringValue(addHeaderPart ctx blocks))) |> ignore)
            h.Even |> Option.iter (fun blocks -> sectPr.AppendChild(Wordprocessing.HeaderReference(Type = EnumValue Wordprocessing.HeaderFooterValues.Even, Id = StringValue(addHeaderPart ctx blocks))) |> ignore; ctx.NeedsEvenAndOddHeaders <- true))

        props.Footer
        |> Option.iter (fun f ->
            f.First |> Option.iter (fun blocks -> sectPr.AppendChild(Wordprocessing.FooterReference(Type = EnumValue Wordprocessing.HeaderFooterValues.First, Id = StringValue(addFooterPart ctx blocks))) |> ignore; titlePage <- true)
            f.Default |> Option.iter (fun blocks -> sectPr.AppendChild(Wordprocessing.FooterReference(Type = EnumValue Wordprocessing.HeaderFooterValues.Default, Id = StringValue(addFooterPart ctx blocks))) |> ignore)
            f.Even |> Option.iter (fun blocks -> sectPr.AppendChild(Wordprocessing.FooterReference(Type = EnumValue Wordprocessing.HeaderFooterValues.Even, Id = StringValue(addFooterPart ctx blocks))) |> ignore; ctx.NeedsEvenAndOddHeaders <- true))

        // `<w:footnotePr>`/`<w:endnotePr>` sit between the header/footer references and
        // `<w:type>` in CT_SectPr's own fixed element order - schema-valid only in that
        // position, same reasoning `<w:type>`'s own note right below gives.
        props.FootnoteNumbering
        |> Option.iter (fun s ->
            let fpr = Wordprocessing.FootnoteProperties()
            fpr.NumberingFormat <- Wordprocessing.NumberingFormat(Val = EnumValue(numberFormatKindToW s.Format))
            s.StartAt |> Option.iter (fun n -> fpr.NumberingStart <- Wordprocessing.NumberingStart(Val = UInt16Value(uint16 n)))
            fpr.NumberingRestart <- Wordprocessing.NumberingRestart(Val = EnumValue(noteNumberRestartToW s.Restart))
            sectPr.AppendChild(fpr) |> ignore)

        props.EndnoteNumbering
        |> Option.iter (fun s ->
            let epr = Wordprocessing.EndnoteProperties()
            epr.NumberingFormat <- Wordprocessing.NumberingFormat(Val = EnumValue(numberFormatKindToW s.Format))
            s.StartAt |> Option.iter (fun n -> epr.NumberingStart <- Wordprocessing.NumberingStart(Val = UInt16Value(uint16 n)))
            epr.NumberingRestart <- Wordprocessing.NumberingRestart(Val = EnumValue(noteNumberRestartToW s.Restart))
            sectPr.AppendChild(epr) |> ignore)

        // `<w:type>` sits between the header/footer references and `<w:pgSz>` in CT_SectPr's
        // own fixed element order - schema-valid only in that position.
        sectionBreakTypeToW props.BreakType
        |> Option.iter (fun v -> sectPr.AppendChild(Wordprocessing.SectionType(Val = EnumValue v)) |> ignore)

        sectPr.AppendChild(pageSizeToW props.PageSize props.Orientation) |> ignore
        sectPr.AppendChild(pageMarginsToW props.Margins) |> ignore

        if props.Columns > 1 then
            sectPr.AppendChild(Wordprocessing.Columns(EqualWidth = OnOffValue true, ColumnCount = Int16Value(int16 props.Columns))) |> ignore

        props.PageNumberStart |> Option.iter (fun n -> sectPr.AppendChild(Wordprocessing.PageNumberType(Start = Int32Value n)) |> ignore)

        if titlePage then
            sectPr.AppendChild(Wordprocessing.TitlePage()) |> ignore

        sectPr

    // --- Document protection ---------------------------------------------------------------

    let private editRestrictionToW (e: EditRestriction) : Wordprocessing.DocumentProtectionValues =
        match e with
        | ReadOnlyRestriction -> Wordprocessing.DocumentProtectionValues.ReadOnly
        | CommentsOnlyRestriction -> Wordprocessing.DocumentProtectionValues.Comments
        | TrackedChangesOnlyRestriction -> Wordprocessing.DocumentProtectionValues.TrackedChanges
        | FormsOnlyRestriction -> Wordprocessing.DocumentProtectionValues.Forms

    /// The modern salted-iterated-SHA512 password scheme (ECMA-376 `legacyPassword`
    /// hashing) - simpler to implement correctly than Excel's classic XOR hash and the
    /// scheme current Word versions themselves default to. Like Excel's own password
    /// fields, this never round-trips back to plaintext (see `Protection.DocumentProtection.
    /// Password`'s own doc comment) - unverified against real Word (no Word available in
    /// this environment to confirm acceptance), same "verify separately" caution Excel gives
    /// its own Sparklines feature.
    let private hashPassword (password: string) : string * string * int =
        let salt = Array.zeroCreate<byte> 16
        System.Security.Cryptography.RandomNumberGenerator.Fill(salt)
        let spinCount = 100000
        use sha = System.Security.Cryptography.SHA512.Create()
        let pwdBytes = System.Text.Encoding.Unicode.GetBytes(password)
        let mutable h = sha.ComputeHash(Array.append salt pwdBytes)

        for i in 0 .. spinCount - 1 do
            let iterBytes = BitConverter.GetBytes(i)
            h <- sha.ComputeHash(Array.append iterBytes h)

        Convert.ToBase64String(h), Convert.ToBase64String(salt), spinCount

    let private documentProtectionToW (dp: DocumentProtection) : Wordprocessing.DocumentProtection option =
        match dp.Edit with
        | None -> None
        | Some edit ->
            let el = Wordprocessing.DocumentProtection(Edit = EnumValue(editRestrictionToW edit), Enforcement = OnOffValue true)

            dp.Password
            |> Option.iter (fun pwd ->
                let hash, salt, spin = hashPassword pwd
                el.CryptographicProviderType <- EnumValue Wordprocessing.CryptProviderValues.RsaAdvancedEncryptionStandard
                el.CryptographicAlgorithmClass <- EnumValue Wordprocessing.CryptAlgorithmClassValues.Hash
                el.CryptographicAlgorithmType <- EnumValue Wordprocessing.CryptAlgorithmValues.TypeAny
                el.CryptographicAlgorithmSid <- Int32Value 14 // SHA-512
                el.CryptographicSpinCount <- UInt32Value(uint32 spin)
                el.HashValue <- Base64BinaryValue hash
                el.SaltValue <- Base64BinaryValue salt)

            Some el

    // --- Top-level orchestration -------------------------------------------------------------

    /// A section's own `SectionProperties` is embedded either as the body's own trailing
    /// `<w:sectPr>` (the last section) or as a `<w:sectPr>` inside the last paragraph of an
    /// earlier section (every other section) - the real WordprocessingML representation of a
    /// "next page" section break. A section whose body doesn't end in a paragraph (e.g. ends
    /// in a table) gets a trailing empty paragraph to carry it, matching what Word itself does.
    let private sectionsToBodyChildren (ctx: Ctx) (sections: Section list) : OpenXmlElement list =
        let n = List.length sections

        sections
        |> List.mapi (fun i sec ->
            let elements = sec.Body |> List.map (blockToW ctx)
            let sectPrEl = sectionPropertiesToW ctx sec.Properties

            if i = n - 1 then
                elements @ [ sectPrEl :> OpenXmlElement ]
            else
                match List.rev elements with
                | (:? Wordprocessing.Paragraph as lastPara) :: restRev ->
                    let pPr =
                        match lastPara.ParagraphProperties with
                        | null ->
                            let p = Wordprocessing.ParagraphProperties()
                            lastPara.PrependChild(p) |> ignore
                            p
                        | existing -> existing

                    pPr.AppendChild(sectPrEl) |> ignore
                    List.rev restRev @ [ lastPara :> OpenXmlElement ]
                | _ ->
                    let trailerPPr = Wordprocessing.ParagraphProperties()
                    trailerPPr.AppendChild(sectPrEl) |> ignore
                    let trailer = Wordprocessing.Paragraph()
                    trailer.AppendChild(trailerPPr) |> ignore
                    elements @ [ trailer :> OpenXmlElement ])
        |> List.concat

    let private docTypeOf (doc: Document) : WordprocessingDocumentType =
        if doc.VbaProject.IsSome then
            WordprocessingDocumentType.MacroEnabledDocument
        else
            WordprocessingDocumentType.Document

    let private writeDocument (doc: Document) (wordDoc: WordprocessingDocument) : unit =
        let mainPart = wordDoc.AddMainDocumentPart()
        let ctx =
            { MainPart = mainPart
              Comments = ResizeArray()
              OpenBookmarkRanges = Collections.Generic.Dictionary()
              OpenCommentRanges = Collections.Generic.Dictionary()
              Footnotes = ResizeArray()
              Endnotes = ResizeArray()
              NextBookmarkId = 1
              NextCommentId = 0
              NextDrawingId = 1u
              NextFootnoteId = 1
              NextEndnoteId = 1
              NextRevisionId = 1
              InsideDeletion = false
              NeedsEvenAndOddHeaders = false }

        let stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>()
        stylesPart.Styles <- stylesToOpenXml doc.Styles
        tableStylesToOpenXml doc.TableStyles |> List.iter (fun s -> stylesPart.Styles.AppendChild(s) |> ignore)

        if not doc.Numbering.IsEmpty then
            let numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>()
            numberingPart.Numbering <- numberingToOpenXml doc.Numbering

        let body = Wordprocessing.Body()
        sectionsToBodyChildren ctx doc.Sections |> List.iter (body.AppendChild >> ignore)

        let protectionEl = doc.Protection |> Option.bind documentProtectionToW

        if protectionEl.IsSome || ctx.NeedsEvenAndOddHeaders then
            let settingsPart = mainPart.AddNewPart<DocumentSettingsPart>()
            let settings = Wordprocessing.Settings()
            protectionEl |> Option.iter (fun el -> settings.AppendChild(el) |> ignore)

            if ctx.NeedsEvenAndOddHeaders then
                settings.AppendChild(Wordprocessing.EvenAndOddHeaders()) |> ignore

            settingsPart.Settings <- settings

        // NOT `Wordprocessing.Document(body)` - F# resolves that to the `IEnumerable<
        // OpenXmlElement>` constructor overload (since `Body` itself implements that
        // interface over ITS OWN children), which re-parents `body`'s children directly
        // under `Document` and throws ("part of a tree") since they're already parented
        // to `body`. Constructing empty then appending sidesteps the ambiguous overload.
        let documentEl = Wordprocessing.Document()
        documentEl.AppendChild(body) |> ignore
        mainPart.Document <- documentEl

        if ctx.Comments.Count > 0 then
            let commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>()
            commentsPart.Comments <- Wordprocessing.Comments(ctx.Comments |> Seq.map (fun c -> c :> OpenXmlElement))

        if ctx.Footnotes.Count > 0 then
            let footnotesPart = mainPart.AddNewPart<FootnotesPart>()
            let footnotesEl = Wordprocessing.Footnotes()

            let separator = Wordprocessing.Footnote(Id = IntegerValue -1, Type = EnumValue Wordprocessing.FootnoteEndnoteValues.Separator)
            separator.AppendChild(separatorParagraph (Wordprocessing.SeparatorMark())) |> ignore
            let continuationSeparator = Wordprocessing.Footnote(Id = IntegerValue 0, Type = EnumValue Wordprocessing.FootnoteEndnoteValues.ContinuationSeparator)
            continuationSeparator.AppendChild(separatorParagraph (Wordprocessing.ContinuationSeparatorMark())) |> ignore
            footnotesEl.AppendChild(separator) |> ignore
            footnotesEl.AppendChild(continuationSeparator) |> ignore

            ctx.Footnotes |> Seq.sortBy fst |> Seq.iter (fun (_, note) -> footnotesEl.AppendChild(note) |> ignore)
            footnotesPart.Footnotes <- footnotesEl

        if ctx.Endnotes.Count > 0 then
            let endnotesPart = mainPart.AddNewPart<EndnotesPart>()
            let endnotesEl = Wordprocessing.Endnotes()

            let separator = Wordprocessing.Endnote(Id = IntegerValue -1, Type = EnumValue Wordprocessing.FootnoteEndnoteValues.Separator)
            separator.AppendChild(separatorParagraph (Wordprocessing.SeparatorMark())) |> ignore
            let continuationSeparator = Wordprocessing.Endnote(Id = IntegerValue 0, Type = EnumValue Wordprocessing.FootnoteEndnoteValues.ContinuationSeparator)
            continuationSeparator.AppendChild(separatorParagraph (Wordprocessing.ContinuationSeparatorMark())) |> ignore
            endnotesEl.AppendChild(separator) |> ignore
            endnotesEl.AppendChild(continuationSeparator) |> ignore

            ctx.Endnotes |> Seq.sortBy fst |> Seq.iter (fun (_, note) -> endnotesEl.AppendChild(note) |> ignore)
            endnotesPart.Endnotes <- endnotesEl

        doc.VbaProject
        |> Option.iter (fun bytes ->
            let vbaPart = mainPart.AddNewPart<VbaProjectPart>()
            use stream = new MemoryStream(bytes)
            vbaPart.FeedData(stream))

        let p = doc.Properties

        // Only touches `docProps/core.xml`/`app.xml` at all when at least one field is
        // set - an all-`None` `DocumentProperties` writes neither part, so it round-trips
        // back to `DocumentProperties.Default` exactly (see that type's own doc comment).
        if p <> DocumentProperties.Default then
            p.Title |> Option.iter (fun v -> wordDoc.PackageProperties.Title <- v)
            p.Author |> Option.iter (fun v -> wordDoc.PackageProperties.Creator <- v)
            p.Subject |> Option.iter (fun v -> wordDoc.PackageProperties.Subject <- v)
            p.Keywords |> Option.iter (fun v -> wordDoc.PackageProperties.Keywords <- v)
            p.Comments |> Option.iter (fun v -> wordDoc.PackageProperties.Description <- v)
            p.Category |> Option.iter (fun v -> wordDoc.PackageProperties.Category <- v)

            p.Company
            |> Option.iter (fun company ->
                let extPart = wordDoc.AddExtendedFilePropertiesPart()
                extPart.Properties <- ExtendedProperties.Properties(Company = ExtendedProperties.Company(company)))

    let saveToStream (doc: Document) (stream: Stream) : unit =
        use wordDoc = WordprocessingDocument.Create(stream, docTypeOf doc)
        writeDocument doc wordDoc

    let saveToFile (doc: Document) (path: string) : unit =
        use wordDoc = WordprocessingDocument.Create(path, docTypeOf doc)
        writeDocument doc wordDoc
