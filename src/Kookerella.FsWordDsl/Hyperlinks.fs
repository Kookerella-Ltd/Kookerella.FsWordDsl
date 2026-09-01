namespace Kookerella.FsWordDsl

/// A hyperlink's destination - external (any URL, including `mailto:`) or internal (a
/// same-document bookmark reference, i.e. `Bookmark.Name` from `Model.fs`'s `Inline.
/// Bookmark`). Unlike Excel's `HyperlinkTarget` (which decorates a cell range that already
/// exists), a Word hyperlink wraps the run(s) it applies to directly - see `Inline.
/// Hyperlink` in `Model.fs`.
[<AutoOpen>]
module Hyperlinks =

    type HyperlinkTarget =
        | ExternalUrl of string
        | InternalBookmark of string
