namespace Kookerella.FsWordDsl

/// Named styles (`styles.xml`) - the one styling concept Excel deliberately doesn't model
/// (see its `MAPPING.md`'s closing note on style interning), but central to how real Word
/// documents are authored: a `Paragraph.StyleId` referencing `"Heading1"` is far more common
/// in practice than direct formatting on every paragraph. `Interpreter/StyleRegistry.fs`
/// ensures every referenced `StyleId` (and its `BasedOn` chain) actually gets written.
[<AutoOpen>]
module NamedStyles =

    type StyleTargetType =
        | ParagraphStyleType
        | CharacterStyleType

    /// `RunFormat`/`ParaFormat` are this style's own contribution - `BasedOn` supplies the
    /// rest via inheritance, the same way Word itself resolves a style's effective
    /// formatting. This DSL does not resolve the inheritance chain itself (that's real
    /// Word's job when it renders/edits the file); `Writer` just needs `BasedOn` to name a
    /// style that's either one of the built-ins below or another entry in the same
    /// `Document.Styles` list, and writes the chain through unmodified.
    type StyleDefinition =
        { Id: string
          Name: string
          Type: StyleTargetType
          BasedOn: string option
          RunFormat: RunStyle option
          ParaFormat: ParagraphFormat option }

    /// A small catalog of the style ids real Word documents reach for constantly, with
    /// explicit formatting given here rather than relying on Word's own built-in template
    /// defaults for these well-known ids - deterministic either way a file is opened, in
    /// keeping with this DSL's "what's written is what's read back" round-trip philosophy.
    /// Not exhaustive - any other `StyleDefinition` a caller builds by hand works the same
    /// way; these are just the common case pre-built.
    module BuiltInStyles =

        let normal: StyleDefinition =
            { Id = "Normal"
              Name = "Normal"
              Type = ParagraphStyleType
              BasedOn = None
              RunFormat = Some { RunStyle.Default with FontFamily = Some "Calibri"; Size = Some 11.0 }
              // Not `Some ParagraphFormat.Default` - an all-defaults format writes as an
              // empty `<w:pPr/>`, which reads back as `None` (nothing there to distinguish
              // it from "no format at all") - `Some ParagraphFormat.Default` and `None` are
              // semantically identical, so this avoids a spurious round-trip mismatch.
              ParaFormat = None }

        let heading1: StyleDefinition =
            { Id = "Heading1"
              Name = "heading 1"
              Type = ParagraphStyleType
              BasedOn = Some "Normal"
              RunFormat =
                Some
                    { RunStyle.Default with
                        Bold = true
                        Size = Some 16.0
                        Color = Some(Rgb(47uy, 84uy, 150uy)) }
              ParaFormat =
                Some
                    { ParagraphFormat.Default with
                        SpacingBefore = Some 12.0
                        SpacingAfter = Some 6.0
                        KeepWithNext = true } }

        let heading2: StyleDefinition =
            { Id = "Heading2"
              Name = "heading 2"
              Type = ParagraphStyleType
              BasedOn = Some "Normal"
              RunFormat =
                Some
                    { RunStyle.Default with
                        Bold = true
                        Size = Some 14.0
                        Color = Some(Rgb(47uy, 84uy, 150uy)) }
              ParaFormat =
                Some
                    { ParagraphFormat.Default with
                        SpacingBefore = Some 10.0
                        SpacingAfter = Some 4.0
                        KeepWithNext = true } }

        let heading3: StyleDefinition =
            { Id = "Heading3"
              Name = "heading 3"
              Type = ParagraphStyleType
              BasedOn = Some "Normal"
              RunFormat =
                Some
                    { RunStyle.Default with
                        Bold = true
                        Size = Some 13.0
                        Color = Some(Rgb(47uy, 84uy, 150uy)) }
              ParaFormat =
                Some
                    { ParagraphFormat.Default with
                        SpacingBefore = Some 8.0
                        SpacingAfter = Some 4.0
                        KeepWithNext = true } }

        let title: StyleDefinition =
            { Id = "Title"
              Name = "Title"
              Type = ParagraphStyleType
              BasedOn = Some "Normal"
              RunFormat = Some { RunStyle.Default with Bold = true; Size = Some 28.0 }
              ParaFormat = Some { ParagraphFormat.Default with SpacingAfter = Some 12.0 } }

        /// The style Word gives every list paragraph by default - just a left indent, no
        /// character formatting of its own.
        let listParagraph: StyleDefinition =
            { Id = "ListParagraph"
              Name = "List Paragraph"
              Type = ParagraphStyleType
              BasedOn = Some "Normal"
              RunFormat = None
              ParaFormat = Some { ParagraphFormat.Default with Indentation = Some { Indentation.None with Left = Some 36.0 } } }

        /// The character style Word applies to a hyperlink's own runs (blue + underlined) -
        /// `SheetDsl.hyperlink`'s Word analog (`DocumentDsl.hyperlink`, see `Builders.fs`)
        /// applies this automatically so callers don't have to restate it on every run.
        let hyperlinkCharStyle: StyleDefinition =
            { Id = "Hyperlink"
              Name = "Hyperlink"
              Type = CharacterStyleType
              BasedOn = None
              RunFormat =
                Some
                    { RunStyle.Default with
                        Color = Some(Rgb(5uy, 99uy, 193uy))
                        Underline = Some SingleUnderline }
              ParaFormat = None }

        /// Every built-in above, for `Document.Styles`' own default when a caller passes
        /// none explicitly - see `Builders.document`.
        let all = [ normal; heading1; heading2; heading3; title; listParagraph; hyperlinkCharStyle ]
