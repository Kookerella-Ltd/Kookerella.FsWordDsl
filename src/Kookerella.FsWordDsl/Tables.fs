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

    /// Cell margins - the table-wide default (`w:tblPr/w:tblCellMar`, see `TableEntry.
    /// CellMargins`) and a single cell's own override (`w:tcPr/w:tcMar`, see
    /// `TableCellProps.Margins`) share this exact shape. All fields in points.
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

    /// Per-cell overrides. `GridSpan` (horizontal merge - "span N grid columns") and
    /// `VerticalMerge` (vertical merge) are independent and can combine on the same cell,
    /// matching real Word. `Shading`/`Borders`/`Width`/`Margins` override the table's own
    /// defaults for just this cell; `None` means "inherit from the table."
    type TableCellProps =
        { GridSpan: int option
          VerticalMerge: VerticalMergeKind option
          Shading: Color option
          Borders: TableBorders option
          /// Points.
          Width: float option
          /// This cell's own margin override (`w:tcPr/w:tcMar`) - see `CellMargins`'s own
          /// doc comment for the table-wide default this overrides.
          Margins: CellMargins option }

        static member Default =
            { GridSpan = None
              VerticalMerge = None
              Shading = None
              Borders = None
              Width = None
              Margins = None }

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
    /// OOXML defines thirteen possible regions (whole table, first/last row, first/last
    /// column, two banding axes - each with an "odd"/"first" and "even"/"second" band -
    /// and four corner cells); `TableStyleDefinition` models eleven of them, leaving only
    /// the "second"/even band of each banding axis undistinguished from `WholeTable`'s own
    /// background (see that type's own doc comment) - same "narrow scope" posture the rest
    /// of this module takes (e.g. `TableBorders` not modeling diagonal cell borders).
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
    /// `"TableGrid"`. `BandedRow`/`BandedColumn` apply to `w:type="band1Horz"`/`"band1Vert"`
    /// (the odd/first band on each axis) only - a distinct look for the even band
    /// (`band2Horz`/`band2Vert`) isn't modeled, since in practice a banded table's "off"
    /// band is just `WholeTable`'s own default background showing through.
    type TableStyleDefinition =
        { Id: string
          Name: string
          BasedOn: string option
          /// The style's own base table borders (`w:tblPr/w:tblBorders`), same shape used
          /// for a `TableEntry`'s direct borders.
          Borders: TableBorders option
          WholeTable: TableStyleRegion
          FirstRow: TableStyleRegion
          LastRow: TableStyleRegion
          FirstColumn: TableStyleRegion
          LastColumn: TableStyleRegion
          BandedRow: TableStyleRegion
          BandedColumn: TableStyleRegion
          /// The four corner cells (`w:type="neCell"`/`"nwCell"`/`"seCell"`/`"swCell"`) -
          /// where a row-band region and a column-band region would otherwise overlap.
          NorthEastCell: TableStyleRegion
          NorthWestCell: TableStyleRegion
          SouthEastCell: TableStyleRegion
          SouthWestCell: TableStyleRegion }

        static member Default =
            { Id = ""
              Name = ""
              BasedOn = None
              Borders = None
              WholeTable = TableStyleRegion.None
              FirstRow = TableStyleRegion.None
              LastRow = TableStyleRegion.None
              FirstColumn = TableStyleRegion.None
              LastColumn = TableStyleRegion.None
              BandedRow = TableStyleRegion.None
              BandedColumn = TableStyleRegion.None
              NorthEastCell = TableStyleRegion.None
              NorthWestCell = TableStyleRegion.None
              SouthEastCell = TableStyleRegion.None
              SouthWestCell = TableStyleRegion.None }

