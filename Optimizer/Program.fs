(*
author: recs
date: 2021-02-05
summary: generate .dxf files from opened documents in application cad solidedge.
*)


open System
open System.IO
open SolidEdgeFramework
open SolidEdgeCommunity.Extensions
open SolidEdgeDraft

[<STAThread>]
[<EntryPoint>]
let main argv =
    try
        SolidEdgeCommunity.OleMessageFilter.Register()
        let application = SolidEdgeCommunity.SolidEdgeUtils.Connect(false)
        let document = application.ActiveDocument
        let draft = document :?> DraftDocument
        let bomQantity = draft.PartsLists.Count

        printfn "Quantity of Part-list = %i" bomQantity

        printfn "-------------------------------------------------------------"
        let maxRowsFirstPage = draft.PartsLists.[1].MaximumRowsFirstPage
        let maxRowsAdditionalPages = draft.PartsLists.[1].MaximumRowsAdditionalPages

        // Rows
        printfn "%30s | %30s |" "RowsFirstPage" "RowsAdditionalPages"
        printfn "%30i | %30i |" maxRowsFirstPage maxRowsAdditionalPages

        let maxHeightFirstPage = draft.PartsLists.[1].MaximumHeightFirstPage
        let maxHeightAdditionalPages = draft.PartsLists.[1].MaximumHeightAdditionalPages

        // Height in meter
        printfn "%30s | %30s |" "HeightFirstPage (meter)" "HeightAdditionalPages (meter)"
        printfn "%30f | %30f |" maxHeightFirstPage maxHeightAdditionalPages

        // Height in meter
        printfn "%30s | %30s |" "HeightFirstPage (inch )" "HeightAdditionalPages (inch )"
        printfn "%30f | %30f |" (maxHeightFirstPage*39.3701) (maxHeightAdditionalPages*39.3701)
        // we could round it 0.000
        0 // exit code

    finally
        SolidEdgeCommunity.OleMessageFilter.Unregister()
        printfn "Press any key to exit"
        Console.ReadKey() |> ignore