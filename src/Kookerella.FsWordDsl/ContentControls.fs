namespace Kookerella.FsWordDsl

open System

/// Structured document tags (`w:sdt`) - Word's "content controls": form fields a template
/// author drops into a document (a labelled text box, a dropdown, a date picker, a
/// checkbox) that real Word constrains editing around. Five kinds are modeled here - the
/// ones actually common in real-world templates - out of Word's full set: picture
/// controls, group/repeating-section controls (structural, rare in hand-authored
/// documents), citation/bibliography controls, XML data binding (`w:dataBinding` - binds a
/// control's value to an external XML part), and placeholder text (`w:placeholder` -
/// references a *separate* "glossary document" part by name rather than carrying its own
/// text inline) are all documented gaps, not covered by any case here.
[<AutoOpen>]
module ContentControls =

    /// Which of Word's control kinds a content control is, plus that kind's own extra
    /// data. `DropDownControl`'s `editable` distinguishes `w:dropDownList` (pick only,
    /// `false`) from `w:comboBox` (pick or type free text, `true`) - both carry the
    /// identical `items` shape on the wire, differing only in that one marker element.
    /// `DateControl`'s `format` is a Word date-format pattern (e.g. `"MM/dd/yyyy"`); the
    /// control's own currently-displayed text still lives in the wrapping `Inline`/`Block`
    /// case's own `content`, the same "cachedResult is just what's currently shown"
    /// posture `Inline.Field` takes - `fullDate`/`format` here are metadata about *how*
    /// that display text was produced, not the text itself.
    type ContentControlType =
        | PlainTextControl of multiLine: bool
        | RichTextControl
        | DropDownControl of items: (string * string) list * editable: bool
        | DateControl of fullDate: DateTime option * format: string option
        /// `Office2010.Word.SdtContentCheckBox` on the wire (`w14:checkbox`) - a different
        /// SDK namespace from every other control kind here, which all live under
        /// `Wordprocessing` - see `Interpreter/Writer.fs`/`Reader.fs`'s own notes on this
        /// case for the qualification this needs. `checkedSymbol`/`uncheckedSymbol` are the
        /// custom checked/unchecked glyphs (`w14:checkedState`/`w14:uncheckedState`) as
        /// `(font, hexCharCode) option` - e.g. `Some("Wingdings", "2612")` for a filled
        /// checkbox glyph - `None` leaves Word's own plain checkmark/empty-box default.
        | CheckBoxControl of checked_: bool * checkedSymbol: (string * string) option * uncheckedSymbol: (string * string) option

    /// Which of Word's `w:lock` values (`w:sdtPr/w:lock`) restrict editing a content
    /// control - `None` on `ContentControlProps.Lock` is Word's own default (`unlocked`,
    /// not written), same "only write what differs from the default" posture the rest of
    /// this DSL takes.
    type ContentControlLock =
        /// `sdtLocked` - the control itself can't be deleted, but its content can still be
        /// edited.
        | LockDeletion
        /// `contentLocked` - the control's content can't be edited, but the control itself
        /// can still be deleted.
        | LockContentEditing
        /// `sdtContentLocked` - neither the control nor its content can be edited.
        | LockDeletionAndContentEditing

    /// `Alias`/`Tag` are Word's own `w:alias` (human-readable title, shown in Word's UI)
    /// and `w:tag` (machine-readable id, for programmatic lookup) - both optional, same as
    /// real Word leaves them when a template author doesn't set them.
    type ContentControlProps =
        { Alias: string option
          Tag: string option
          Lock: ContentControlLock option
          Type: ContentControlType }
