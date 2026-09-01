namespace Kookerella.FsWordDsl.Interpreter

open System.Text
open Kookerella.FsWordDsl

/// `Document` -> F# *source text*: renders a value back out as a self-contained `.fsx`
/// script that rebuilds an equivalent file when run - a code-generating counterpart to
/// `Reader`, same idea as Excel's own `CodeGen`. Unlike Excel's, this always renders every
/// field explicitly (no "only mention what differs from Default" diffing) - simpler and
/// still correct, just more verbose output; a good target for a future pass once this DSL
/// has more real-world mileage.
module CodeGen =

    let private quote (s: string) : string =
        let escaped = s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "")
        "\"" + escaped + "\""

    let private renderOption (render: 'a -> string) (o: 'a option) : string =
        match o with
        | None -> "None"
        // The inner value is ALSO parenthesized, not just the whole `Some ...` - a DU case
        // application like `Rgb(47uy, 84uy, 150uy)` needs its own parens as `Some`'s
        // argument (`Some Rgb(47uy, 84uy, 150uy)` parses as `(Some Rgb) (47uy, 84uy, 150uy)`,
        // not `Some (Rgb(...))`) - redundant-but-harmless for a record/tuple literal, which
        // don't need it, so this always adds them rather than trying to tell the two apart.
        | Some v -> sprintf "(Some (%s))" (render v)

    let private renderList (render: 'a -> string) (xs: 'a list) : string =
        "[ " + (xs |> List.map render |> String.concat "; ") + " ]"

    let private renderTuple2 (renderA: 'a -> string) (renderB: 'b -> string) ((a, b): 'a * 'b) : string =
        sprintf "(%s, %s)" (renderA a) (renderB b)

    let private renderColor (c: Color) : string =
        match c with
        | Rgb(r, g, b) -> sprintf "Rgb(%duy, %duy, %duy)" r g b
        | Auto -> "Auto"

    let private renderHighlight (h: HighlightColor) : string = sprintf "%A" h
    let private renderUnderline (u: UnderlineStyle) : string =
        match u with
        | OtherUnderline raw -> sprintf "OtherUnderline %s" (quote raw)
        | other -> sprintf "%A" other

    let private renderVerticalPosition (v: VerticalPosition) : string = sprintf "%A" v

    let private renderRunStyle (s: RunStyle) : string =
        sprintf
            "{ RunStyle.Default with FontFamily = %s; Size = %s; Bold = %b; Italic = %b; Underline = %s; Strikethrough = %b; Color = %s; Highlight = %s; VerticalPosition = %s }"
            (renderOption quote s.FontFamily)
            (renderOption string s.Size)
            s.Bold
            s.Italic
            (renderOption renderUnderline s.Underline)
            s.Strikethrough
            (renderOption renderColor s.Color)
            (renderOption renderHighlight s.Highlight)
            (renderOption renderVerticalPosition s.VerticalPosition)

    let private renderIndentation (i: Indentation) : string =
        sprintf
            "{ Left = %s; Right = %s; FirstLine = %s; Hanging = %s }"
            (renderOption string i.Left)
            (renderOption string i.Right)
            (renderOption string i.FirstLine)
            (renderOption string i.Hanging)

    let private renderLineSpacing (ls: LineSpacingRule) : string =
        match ls with
        | AtLeastSpacing p -> sprintf "AtLeastSpacing %g" p
        | ExactlySpacing p -> sprintf "ExactlySpacing %g" p
        | MultipleSpacing f -> sprintf "MultipleSpacing %g" f
        | other -> sprintf "%A" other

    let private renderParagraphFormat (f: ParagraphFormat) : string =
        sprintf
            "{ ParagraphFormat.Default with Alignment = %s; SpacingBefore = %s; SpacingAfter = %s; LineSpacing = %s; Indentation = %s; KeepWithNext = %b; PageBreakBefore = %b }"
            (renderOption (sprintf "%A") f.Alignment)
            (renderOption string f.SpacingBefore)
            (renderOption string f.SpacingAfter)
            (renderOption renderLineSpacing f.LineSpacing)
            (renderOption renderIndentation f.Indentation)
            f.KeepWithNext
            f.PageBreakBefore

    let private renderBorderSide (s: BorderSide) : string =
        let style =
            match s.Style with
            | OtherLine raw -> sprintf "OtherLine %s" (quote raw)
            | other -> sprintf "%A" other

        sprintf "{ Style = %s; Width = %s; Color = %s }" style (renderOption string s.Width) (renderOption renderColor s.Color)

    let private renderBorderStyle (b: BorderStyle) : string =
        sprintf
            "{ Left = %s; Right = %s; Top = %s; Bottom = %s }"
            (renderOption renderBorderSide b.Left)
            (renderOption renderBorderSide b.Right)
            (renderOption renderBorderSide b.Top)
            (renderOption renderBorderSide b.Bottom)

    let private renderTableBorders (b: TableBorders) : string =
        sprintf
            "{ Outer = %s; InsideHorizontal = %s; InsideVertical = %s }"
            (renderBorderStyle b.Outer)
            (renderOption renderBorderSide b.InsideHorizontal)
            (renderOption renderBorderSide b.InsideVertical)

    let private renderTableStyleRef (s: TableStyleRef) : string =
        sprintf
            "{ Name = %s; FirstRowBanding = %b; LastRowBanding = %b; BandedRows = %b; BandedColumns = %b }"
            (quote s.Name)
            s.FirstRowBanding
            s.LastRowBanding
            s.BandedRows
            s.BandedColumns

    let private renderImageEntry (img: ImageEntry) : string =
        // Image bytes are rendered as a base64 literal decoded at script-run time, rather
        // than an unreadable byte-array literal - `System.Convert.FromBase64String` is
        // already open via `System`.
        sprintf
            "{ Data = System.Convert.FromBase64String(%s); Format = %A; WidthEmu = %dL; HeightEmu = %dL; AltText = %s }"
            (quote (System.Convert.ToBase64String(img.Data)))
            img.Format
            img.WidthEmu
            img.HeightEmu
            (renderOption quote img.AltText)

    let private renderHyperlinkTarget (t: HyperlinkTarget) : string =
        match t with
        | ExternalUrl u -> sprintf "ExternalUrl %s" (quote u)
        | InternalBookmark n -> sprintf "InternalBookmark %s" (quote n)

    let rec private renderInline (i: Inline) : string =
        match i with
        | Run(text, style, styleId) -> sprintf "Run(%s, %s, %s)" (quote text) (renderOption renderRunStyle style) (renderOption quote styleId)
        | LineBreak -> "LineBreak"
        | Tab -> "Tab"
        | PageBreak -> "PageBreak"
        | Image img -> sprintf "Image(%s)" (renderImageEntry img)
        | Hyperlink(target, runs, tooltip) ->
            sprintf "Hyperlink(%s, %s, %s)" (renderHyperlinkTarget target) (renderList renderInline runs) (renderOption quote tooltip)
        | Bookmark(name, content) -> sprintf "Bookmark(%s, %s)" (quote name) (renderList renderInline content)
        | Comment(author, initials, date, text, content) ->
            let dateStr = renderOption (fun (d: System.DateTime) -> sprintf "System.DateTime.Parse(%s)" (quote (d.ToString("o")))) date
            sprintf "Comment(%s, %s, %s, %s, %s)" (quote author) (renderOption quote initials) dateStr (quote text) (renderList renderInline content)
        | Field(instr, cached) -> sprintf "Field(%s, %s)" (quote instr) (renderOption quote cached)

    let private renderParagraph (p: Paragraph) : string =
        sprintf
            "{ Inlines = %s; StyleId = %s; Format = %s; Numbering = %s }"
            (renderList renderInline p.Inlines)
            (renderOption quote p.StyleId)
            (renderOption renderParagraphFormat p.Format)
            (renderOption (renderTuple2 string string) p.Numbering)

    let rec private renderBlock (b: Block) : string =
        match b with
        | ParagraphBlock p -> sprintf "ParagraphBlock(%s)" (renderParagraph p)
        | TableBlock t -> sprintf "TableBlock(%s)" (renderTableEntry t)

    and private renderTableCellProps (p: TableCellProps) : string =
        sprintf
            "{ GridSpan = %s; VerticalMerge = %s; Shading = %s; Borders = %s; Width = %s }"
            (renderOption string p.GridSpan)
            (renderOption (sprintf "%A") p.VerticalMerge)
            (renderOption renderColor p.Shading)
            (renderOption renderTableBorders p.Borders)
            (renderOption string p.Width)

    and private renderTableCell (c: TableCell) : string =
        sprintf "{ Content = %s; Props = %s }" (renderList renderBlock c.Content) (renderTableCellProps c.Props)

    and private renderTableRow (r: TableRow) : string =
        sprintf "{ Cells = %s; Height = %s }" (renderList renderTableCell r.Cells) (renderOption string r.Height)

    and private renderTableEntry (t: TableEntry) : string =
        sprintf
            "{ Rows = %s; ColumnWidths = %s; Style = %s; Borders = %s }"
            (renderList renderTableRow t.Rows)
            (renderList string t.ColumnWidths)
            (renderOption renderTableStyleRef t.Style)
            (renderOption renderTableBorders t.Borders)

    let private renderPageSize (p: PageSize) : string =
        match p with
        | OtherPageSize code -> sprintf "OtherPageSize %d" code
        | CustomPageSize(w, h) -> sprintf "CustomPageSize(%g, %g)" w h
        | other -> sprintf "%A" other

    let private renderPageMargins (m: PageMargins) : string =
        sprintf
            "{ Top = %g; Bottom = %g; Left = %g; Right = %g; Header = %g; Footer = %g; Gutter = %g }"
            m.Top
            m.Bottom
            m.Left
            m.Right
            m.Header
            m.Footer
            m.Gutter

    let private renderHeaderFooterSet (h: HeaderFooterSet) : string =
        sprintf
            "{ Default = %s; First = %s; Even = %s }"
            (renderOption (renderList renderBlock) h.Default)
            (renderOption (renderList renderBlock) h.First)
            (renderOption (renderList renderBlock) h.Even)

    let private renderSectionProperties (s: SectionProperties) : string =
        sprintf
            "{ PageSize = %s; Orientation = %A; Margins = %s; Header = %s; Footer = %s; PageNumberStart = %s; Columns = %d }"
            (renderPageSize s.PageSize)
            s.Orientation
            (renderPageMargins s.Margins)
            (renderOption renderHeaderFooterSet s.Header)
            (renderOption renderHeaderFooterSet s.Footer)
            (renderOption string s.PageNumberStart)
            s.Columns

    let private renderSection (s: Section) : string =
        sprintf "{ Body = %s; Properties = %s }" (renderList renderBlock s.Body) (renderSectionProperties s.Properties)

    let private renderStyleDefinition (d: StyleDefinition) : string =
        sprintf
            "{ Id = %s; Name = %s; Type = %A; BasedOn = %s; RunFormat = %s; ParaFormat = %s }"
            (quote d.Id)
            (quote d.Name)
            d.Type
            (renderOption quote d.BasedOn)
            (renderOption renderRunStyle d.RunFormat)
            (renderOption renderParagraphFormat d.ParaFormat)

    let private renderNumberFormatKind (k: NumberFormatKind) : string =
        match k with
        | BulletFormat(glyph, font) -> sprintf "BulletFormat(char %d, %s)" (int glyph) (quote font)
        | OtherFormat raw -> sprintf "OtherFormat %s" (quote raw)
        | other -> sprintf "%A" other

    let private renderListLevel (l: ListLevel) : string =
        sprintf
            "{ Format = %s; Text = %s; IndentLeft = %s; HangingIndent = %s; StartAt = %s }"
            (renderNumberFormatKind l.Format)
            (quote l.Text)
            (renderOption string l.IndentLeft)
            (renderOption string l.HangingIndent)
            (renderOption string l.StartAt)

    let private renderNumberingDefinition (d: NumberingDefinition) : string =
        sprintf "{ Id = %d; Levels = %s }" d.Id (renderList renderListLevel d.Levels)

    let private renderDocumentProtection (p: DocumentProtection) : string =
        sprintf "{ Edit = %s; Password = %s }" (renderOption (sprintf "%A") p.Edit) (renderOption quote p.Password)

    /// Renders `doc` as a self-contained `.fsx` script - see `Api.Document.generateScript`.
    let generate (referenceLines: string list) (outputFileName: string) (doc: Document) : string =
        let sb = StringBuilder()

        for line in referenceLines do
            sb.AppendLine(line) |> ignore

        sb.AppendLine() |> ignore
        sb.AppendLine("open Kookerella.FsWordDsl") |> ignore
        sb.AppendLine() |> ignore

        sb.AppendLine(sprintf "let doc: Document = { Sections = %s; Styles = %s; Numbering = %s; Protection = %s; VbaProject = %s }"
                          (renderList renderSection doc.Sections)
                          (renderList renderStyleDefinition doc.Styles)
                          (renderList renderNumberingDefinition doc.Numbering)
                          (renderOption renderDocumentProtection doc.Protection)
                          (renderOption (fun (b: byte[]) -> sprintf "System.Convert.FromBase64String(%s)" (quote (System.Convert.ToBase64String(b)))) doc.VbaProject))
        |> ignore

        sb.AppendLine() |> ignore
        sb.AppendLine(sprintf "doc |> Document.save %s" (quote outputFileName)) |> ignore

        sb.ToString()
