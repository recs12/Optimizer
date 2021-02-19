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
        let draft = application.ActiveDocument
        let d = draft :?> DraftDocument
        let a = d.PartsLists.Count
        printfn "%i" a
        0 // exit code

    finally
        SolidEdgeCommunity.OleMessageFilter.Unregister()
        printfn "Press any key to exit"
        Console.ReadKey() |> ignore