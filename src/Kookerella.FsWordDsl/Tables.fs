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

    /// A reference to a table style *by name* - either a built-in like `"TableGrid"`, or a
    /// custom one defined in `Document.TableStyles` (see `TableStyleDefinition` below,
    /// which this DSL does now model, narrowly - see that type's own doc comment for what's
    /// covered and what isn't).
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

    /// One conditional-formatting region within a custom table style (`w:tblStylePr`).
    /// OOXML defines ten possible regions (whole table, first/last row, first/last column,
    /// two banding axes, four corner cells); this DSL models only the three a real custom
    /// table style overwhelmingly actually uses in practice - whole-table defaults, a
    /// distinct header row, and alternating row banding - documented as a gap, same "narrow
    /// scope" posture the rest of this module takes (e.g. `TableBorders` not modeling
    /// diagonal cell borders).
    type TableStyleRegion =
        { RunFormat: RunStyle option
          ParaFormat: ParagraphFormat option
          /// This region's own cell background (`w:tcPr/w:shd`) - independent of
          /// `RunFormat`/`ParaFormat`, which cover text/paragraph appearance only.
          CellShading: Color option }

        static member None =
            { RunFormat = None
              ParaFormat = None
              CellShading = None }

    /// A custom table style *definition* (`w:style` with `w:type="table"`, `styles.xml`) -
    /// unlike `TableStyleRef`, which only references a style by name, this actually defines
    /// one. Add it to `Document.TableStyles` (see `Builders.withTableStyles`) and reference
    /// its `Id` from `TableStyleRef.Name` the same way you'd reference a built-in name like
    /// `"TableGrid"`. `BandedRow` applies to `w:type="band1Horz"` (the odd/first band) only -
    /// a distinct look for the even band (`band2Horz`) isn't modeled, since in practice a
    /// banded table's "off" band is just `WholeTable`'s own default background showing
    /// through.
    type TableStyleDefinition =
        { Id: string
          Name: string
          BasedOn: string option
          /// The style's own base table borders (`w:tblPr/w:tblBorders`), same shape used
          /// for a `TableEntry`'s direct borders.
          Borders: TableBorders option
          WholeTable: TableStyleRegion
          FirstRow: TableStyleRegion
          BandedRow: TableStyleRegion }

        static member Default =
            { Id = ""
              Name = ""
              BasedOn = None
              Borders = None
              WholeTable = TableStyleRegion.None
              FirstRow = TableStyleRegion.None
              BandedRow = TableStyleRegion.None }

    /// Table-wide default cell margins (`w:tblPr/w:tblCellMar`) - every cell inherits these
    /// unless it has its own `w:tcMar` override, which this DSL doesn't model per-cell (same
    /// "narrow scope" posture `TableStyleRegion` documents above); a caller who needs one
    /// cell's margins to differ from the table's default is a documented gap. All fields in
    /// points.
    type CellMargins =
        { Top: float option
          Bottom: float option
          Left: float option
          Right: float option }

        static member Default =
            { Top = None
              Bottom = None
              Left = None
              Right = None }
