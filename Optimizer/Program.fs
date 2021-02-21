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
        
        printfn "%i" bomQantity

        let maxRowsPage1 = draft.PartsLists.[1].MaximumRowsFirstPage
        let maxRowsPage2 = draft.PartsLists.[1].MaximumRowsAdditionalPages

        // back Bom data
        // in the layout order
        printfn "%s | %s |" "page2" "page1"
        printfn "%i | %i |" maxRowsPage2 maxRowsPage1

        let maxHeigghtPage1 = draft.PartsLists.[1].MaximumHeightFirstPage
        let maxHeightPage2 = draft.PartsLists.[1].MaximumHeightAdditionalPages

        // front Bom data
        // in the layout order
        printfn "%s | %s |" "page2" "page1"
        printfn "%f | %f |" maxHeightPage2 maxHeigghtPage1


        0 // exit code

    finally
        SolidEdgeCommunity.OleMessageFilter.Unregister()
        printfn "Press any key to exit"
        Console.ReadKey() |> ignore