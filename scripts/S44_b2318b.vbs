' S44 body : 2318 H0 (B). ATP-07 sum check (item11 -> same material as item2), then HIST-07/02: change item1 + item11 to
'            in-stock materials (candidates tried until dbl-click shows no ATP error) -> CREA_ORD -> Z33SDT644 (2 rows) -> re-query
Const ID_2318 = "0407e0e8-8434-11f1-9bbb-000c298aeffe"
Dim r, msg, vb, k, cand, cands, ok11
If Not OpenByID(ID_2318) Then log.WriteLine "ABORT 2318" : Finish : WScript.Quit 10
log.WriteLine "2318 start: " & HState(0)
msg = DblHeader(0)
log.WriteLine "  dblclick: " & msg & " items=" & ItemRows() & " item11=" & It(10,"MATNR") & " item2=" & It(1,"MATNR")

' ---- ATP-07 : item 11 -> PNGB3500013L008 (same as item 2, qty 14+28) ----
r = ChangeMat(10, "PNGB3500013L008")
log.WriteLine "  ATP-07 change item11: " & r & " -> " & It(10,"MATNR")
msg = DblHeader(0)
Shot "ATP_07_2318_sum_dblclick.png"
r = HeaderAction(0, "CREA_ORD", "ATP_07_2318_crea_ord_log.png")
vb = "?"
If OpenByID(ID_2318) Then vb = Hr(0,"VBELN") : r = r & " after: " & HState(0)
Result "ATP-07 same material in 2 items (14+28) -> summed ATP check blocks CREA_ORD", (vb = ""), "shortage (summed) blocks; both rows red in shot", "dblclick " & msg & " | " & r

' ---- HIST-07 prep: change item1 and item11 to in-stock materials ----
If vb = "" Then
    msg = DblHeader(0)
    r = ChangeMat(0, "PNGB3100016L004")
    log.WriteLine "  HIST-07 change item1 -> PNGB3100016L004: " & r & " -> " & It(0,"MATNR")
    cands = Array("PNGB3100013L003", "PNGB2100013L001", "PNGB2100016L001", "PNGB3520024L002", "PNGB3500016L008")
    ok11 = False
    For Each cand In cands
        r = ChangeMat(10, cand)
        msg = DblHeader(0)
        log.WriteLine "  try item11 -> " & cand & " : " & r & " ; dblclick " & msg
        If InStr(msg, "ATP") = 0 And InStr(msg, "E:") = 0 Then ok11 = True : Exit For
    Next
    Shot "HIST_07_2318_after_changes.png"
    log.WriteLine "  item1=" & It(0,"MATNR") & " item11=" & It(10,"MATNR") & " atpClear=" & ok11
    r = HeaderAction(0, "CREA_ORD", "HIST_02_2318_crea_ord_log.png")
    vb = ""
    If OpenByID(ID_2318) Then vb = Hr(0,"VBELN") : r = r & " requery: " & HState(0)
    Shot "HIST_02_2318_after_crea_ord.png"
    Result "HIST-01/02 2318 CREA_ORD after 2 material changes", (vb <> ""), "sales order created", r
    If vb <> "" Then
        msg = DblHeader(0)
        Shot "HIST_03_2318_requery_items.png"
        Result "HIST-03 2318 re-query shows changed materials", (It(0,"MATNR") = "PNGB3100016L004"), "item1=PNGB3100016L004, item11=changed", "item1=" & It(0,"MATNR") & " item11=" & It(10,"MATNR")
        Se16n "Z33SDT644", "Uniq", ID_2318, "SEQNO,MATNR_OLD,MATNR,VBELN,ERNAM,ERDAT,ERZET", "HIST-07 Z33SDT644 2318", 12, "HIST_07_2318_z33sdt644.png"
        ' HIST-04 / HIST-05 : search by original / changed material
        Tcode "/nZSD600"
        If WaitSel() Then
            With session
                .findById("wnd[0]/usr/radP_RBLIST").Select
                .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
                .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = "2318"
                .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = "PNGA0023FD009"
                .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
                .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
                .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
                .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
                .findById("wnd[0]/tbar[1]/btn[8]").press
            End With
            WScript.Sleep 3500
            Dim n4 : n4 = 0
            If session.Info.ScreenNumber = 100 Then n4 = HeaderRows()
            Shot "HIST_04_2318_search_old.png"
            Result "HIST-04 search by original material PNGA0023FD009 finds 2318", (n4 >= 1), "header found", "rows=" & n4 & " sbar=" & Sbar()
        End If
        Tcode "/nZSD600"
        If WaitSel() Then
            With session
                .findById("wnd[0]/usr/radP_RBLIST").Select
                .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = "2318"
                .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = "PNGB3100016L004"
                .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
                .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
                .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
                .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
                .findById("wnd[0]/tbar[1]/btn[8]").press
            End With
            WScript.Sleep 3500
            Dim n5 : n5 = 0
            If session.Info.ScreenNumber = 100 Then n5 = HeaderRows() : msg = DblHeader(0)
            Shot "HIST_05_2318_search_new.png"
            Result "HIST-05 search by changed material PNGB3100016L004 finds 2318 and shows it", (n5 >= 1 And It(0,"MATNR") = "PNGB3100016L004"), "header found, item shows changed material", "rows=" & n5 & " item1=" & It(0,"MATNR")
        End If
    End If
End If
Finish
