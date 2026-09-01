namespace Kookerella.FsWordDsl

/// Table formatting types that don't themselves need to hold arbitrary document content -
/// `TableCell`/`TableRow`/`TableEntry` do (a cell's content is a `Block list`, and `Block`
/// has a table case), so those live in `Model.fs` instead, declared together with `Block`
/// via a mutually-recursive `and` chain rather than forcing an artificial split. This file
/// holds everything about a table's *appearance* that doesn't participate in that recursion.
[<AutoOpen>]
module Tables =

    /// Reuses `Styles.BorderStyle`'s four sides and adds the two "inside gridline" sides a
    /// multi-row/column table has that a single bordered box doesn't.
    type TableBorders =
        { Outer: BorderStyle
          InsideHorizontal: BorderSide option
          InsideVertical: BorderSide option }

        static member None =
            { Outer = BorderStyle.None
              InsideHorizontal = None
              InsideVertical = None }

    /// A cell's vertical merge state - `Restart` begins a new merged group (this cell is
    /// the visible one, spanning downward), `Continue` merges into the group started by the
    /// nearest `Restart` above it in the same column (matching OOXML's own `w:vMerge`
    /// semantics: a bare `Continue` with no properties of its own is what "merged into the
    /// cell above" actually means on the wire).
    type VerticalMergeKind =
        | RestartMerge
        | ContinueMerge

    /// Per-cell overrides. `GridSpan` (horizontal merge - "span N grid columns") and
    /// `VerticalMerge` (vertical merge) are independent and can combine on the same cell,
    /// matching real Word. `Shading`/`Borders`/`Width` override the table's own defaults for
    /// just this cell; `None` means "inherit from the table."
    type TableCellProps =
        { GridSpan: int option
          VerticalMerge: VerticalMergeKind option
          Shading: Color option
          Borders: TableBorders option
          /// Points.
          Width: float option }

        static member Default =
            { GridSpan = None
              VerticalMerge = None
              Shading = None
              Borders = None
              Width = None }

    /// A reference to a table style *by name* (a built-in like `"TableGrid"`, or a custom
    /// one defined elsewhere in the document) - this DSL doesn't model custom table style
    /// *definitions* themselves, only the reference, same documented gap Excel's own
    /// `TableStyle.Name` has for `tableStyles`/`dxf`-based custom styles.
    type TableStyleRef =
        { Name: string
          FirstRowBanding: bool
          LastRowBanding: bool
          BandedRows: bool
          BandedColumns: bool }

        static member Default =
            { Name = "TableGrid"
              FirstRowBanding = false
              LastRowBanding = false
              BandedRows = false
              BandedColumns = false }
