' S42 body : 2318 H0 -> CLOS_CANC -> ATP-02 (dbl-click shortage + CREA_ORD blocked, per-item log)
Const ID_2318 = "0407e0e8-8434-11f1-9bbb-000c298aeffe"
Dim r, msg, vb
If Not OpenByID(ID_2318) Then log.WriteLine "ABORT 2318" : Finish : WScript.Quit 10
log.WriteLine "2318 H0 start: " & HState(0)
If Hr(0,"STATU") = "C" Then
    r = HeaderAction(0, "CLOS_CANC", "B01_2318_clos_canc_log.png")
    If OpenByID(ID_2318) Then r = r & " after: " & HState(0)
    Result "PRE 2318 CLOS_CANC -> STATU B", (Hr(0,"STATU") <> "C"), "STATU B", r
End If
msg = DblHeader(0)
Shot "ATP_02_2318_dblclick.png"
Note "ATP-02 2318 dbl-click", msg & " items=" & ItemRows()
Dim k
For k = 0 To ItemRows() - 1
    log.WriteLine "   I" & k & " SEQNO=" & It(k,"SEQNO") & " MATNR=" & It(k,"MATNR") & " ZMENG=" & It(k,"ZMENG") & " MEINS=" & It(k,"MEINS") & " LGORT=" & It(k,"LGORT")
Next
r = HeaderAction(0, "CREA_ORD", "ATP_02_2318_crea_ord_log.png")
vb = "?"
If OpenByID(ID_2318) Then vb = Hr(0,"VBELN") : r = r & " after: " & HState(0)
Shot "ATP_02_2318_after.png"
Result "ATP-02 2318 (STATU B) ATP shortage blocks CREA_ORD", (vb = ""), "no sales order; per-item shortage in log", r
Finish
