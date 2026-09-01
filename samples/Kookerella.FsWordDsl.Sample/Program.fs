open System
open Kookerella.FsWordDsl
open type Kookerella.FsWordDsl.DocumentDsl

let private cell (text: string) = tableCell ([ para [ run text ] ], props = { TableCellProps.Default with Width = Some 150.0 })

let private headerCell (text: string) =
    tableCell (
        [ para [ run (text, style = { RunStyle.Default with Bold = true; Color = Some Color.white }) ] ],
        props = { TableCellProps.Default with Shading = Some(Rgb(68uy, 84uy, 106uy)); Width = Some 150.0 }
    )

let buildReport () : Document =
    let body =
        [ para ([ run "Quarterly Report" ], styleId = "Title")
          para ([ run "Summary" ], styleId = "Heading1")
          para
              [ run "This report covers "
                run ("Q1 2026", style = { RunStyle.Default with Bold = true })
                run ", with figures for each product line below. See the "
                hyperlink ("full dataset", ExternalUrl "https://github.com/Kookerella-Ltd")
                run " for details." ]
          table (
              [ tableRow [ headerCell "Item"; headerCell "Revenue" ]
                tableRow [ cell "Widgets"; cell "$4,200" ]
                tableRow [ cell "Gadgets"; cell "$1,980" ] ],
              [ 150.0; 150.0 ],
              style = TableStyleRef.Default
          )
          para ([ run "Next steps" ], styleId = "Heading1")
          para ([ run "Finalize the Q2 forecast" ], numbering = (1, 0))
          para ([ run "Review pricing for Gadgets" ], numbering = (1, 0)) ]

    document [ section body ] |> withNumbering [ bulletListDef 1 ]

[<EntryPoint>]
let main argv =
    let path =
        match argv with
        | [| p |] -> p
        | _ -> IO.Path.Combine(IO.Path.GetTempPath(), "fswordsl-sample.docx")

    let doc = buildReport ()
    Document.save path doc
    printfn "Wrote %s" path

    // Reverse transform: read the file we just wrote back into the DSL.
    let roundTripped = Document.load path
    let paragraphCount = roundTripped.Sections |> List.sumBy (fun s -> s.Body |> List.filter (function ParagraphBlock _ -> true | _ -> false) |> List.length)
    let tableCount = roundTripped.Sections |> List.sumBy (fun s -> s.Body |> List.filter (function TableBlock _ -> true | _ -> false) |> List.length)
    printfn "Read back %d section(s): %d paragraph(s), %d table(s)" roundTripped.Sections.Length paragraphCount tableCount

    0
