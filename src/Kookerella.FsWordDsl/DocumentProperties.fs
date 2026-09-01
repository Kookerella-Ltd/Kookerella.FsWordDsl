namespace Kookerella.FsWordDsl

/// Document-level metadata (`docProps/core.xml` and `docProps/app.xml`) - not content, but
/// the properties Word's own File > Info panel shows, and what a search index/file browser
/// reads to describe the document without opening it. Entirely absent from Excel's own
/// `MAPPING.md` scope, but a real gap for anything meant to look like a properly-authored
/// document rather than just content: a generated report with no title/author metadata is
/// a giveaway in a way the body text alone isn't.
[<AutoOpen>]
module DocumentProperties =

    /// All fields optional and `None` by default - `Writer` only touches `docProps/core.xml`/
    /// `app.xml` at all when at least one field is set (see `Api.Document.save`'s own note),
    /// so a document with no properties set round-trips back to `DocumentProperties.Default`
    /// exactly, the same "an all-default value and 'nothing here' must read back identically"
    /// discipline `NamedStyles.BuiltInStyles.normal`'s own doc comment explains the reasoning
    /// for. `Company` is the one field that lives in `app.xml` rather than `core.xml` - a
    /// caller doesn't need to know or care about that split.
    type DocumentProperties =
        { Title: string option
          Author: string option
          Subject: string option
          Keywords: string option
          /// Word's own UI now calls this field "Comments" (File > Info > Properties) even
          /// though OOXML's own package-level name for it is `dc:description`.
          Comments: string option
          Category: string option
          Company: string option }

        static member Default =
            { Title = None
              Author = None
              Subject = None
              Keywords = None
              Comments = None
              Category = None
              Company = None }
