namespace Kookerella.FsWordDsl

/// Document-level editing restrictions (`w:documentProtection`) - Word has no per-section
/// equivalent of Excel's `SheetProtection`, only this one document-wide setting, so unlike
/// the Excel repo (which has both `Protection.fs` types), there's just the one type here.
[<AutoOpen>]
module Protection =

    /// Which single kind of edit Word still allows while the document is protected - these
    /// are mutually exclusive in real Word (protecting for one restricts everything else),
    /// matching OOXML's own `w:documentProtection/@w:edit` enumeration.
    type EditRestriction =
        | ReadOnlyRestriction
        | CommentsOnlyRestriction
        | TrackedChangesOnlyRestriction
        | FormsOnlyRestriction

    /// `Password` is hashed with the same legacy XOR algorithm Excel's `SheetProtection`/
    /// `WorkbookProtection` use for broad compatibility, and never round-trips back to
    /// plaintext (the hash isn't reversible) - same deliberate consequence, not an
    /// oversight, as Excel's own `MAPPING.md` documents for its password fields.
    type DocumentProtection =
        { Edit: EditRestriction option
          Password: string option }

        static member Default = { Edit = None; Password = None }
