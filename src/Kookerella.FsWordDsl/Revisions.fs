namespace Kookerella.FsWordDsl

open System

/// Track changes (`w:ins`/`w:del`) - narrowly scoped to the case that actually matters for
/// almost every real redlined document: visible inserted/deleted content, attributed and
/// dated. Deliberately not modeled (see MAPPING.md for the full list): formatting-change
/// history (`w:rPrChange`/`w:pPrChange` - the *previous* formatting a run/paragraph had
/// before an edit), moves (`w:moveFrom`/`w:moveTo` - Word's "this looks like cut+paste"
/// detection, which without this DSL modeling it just shows as an ordinary delete-then-
/// insert instead, still correct information, just not the special annotation), and
/// table row/cell-level insertion/deletion tracking.
[<AutoOpen>]
module Revisions =

    type RevisionKind =
        | Inserted
        | Deleted

    /// `Date = None` is written as "now" at write time, same convention `Model.Inline.
    /// Comment`'s own `Date` uses and for the same reason - Word records a revision's own
    /// timestamp, this DSL doesn't require the caller to supply one.
    type Revision =
        { Kind: RevisionKind
          Author: string
          Date: DateTime option }
