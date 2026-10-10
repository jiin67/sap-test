' S45 body : parts/boiler delivery+GI flow.
'   1722 (parts, KTGRM Y2): CLOS_CANC -> CREA_ORD -> CLOS_ORD -> CRE_DELIVERY (expect delivery + GI: DIV-02, ICON-03, ALV-04/05) -> VBUP
'   2392 (boiler Y1):        CLOS_CANC -> CREA_ORD -> CLOS_ORD -> (delivery in mixed run)
'   2827 (B, LGORT blank):   CREA_ORD -> CLOS_ORD -> (delivery in mixed run, expect error)
'   DEL-03: 2766 (no SO) -> CRE_DELIVERY blocked
'   DEL-06/REG-04: 2392 + 2827 + 1757(dup) selected together -> CRE_DELIVERY
'   ALV-06: re-query 1722 after GI ; ALV-08 layout popups
Const ID_1722 = "04a27fe5-6007-11f1-9ba0-000c298aeffe"
Const ID_2392 = "04a5f20c-8a85-11f1-9bc0-000c298aeffe"
Const ID_2827 = "04bf872f-a20b-11f1-9bce-000c298aeffe"
Const ID_2766 = "04ad1bf8-a070-11f1-9bcd-000c298aeffe"
Dim r, msg, vb, vl

Function PrepSO(idv, tag)
    ' returns VBELN after CLOS_CANC(if C) -> CREA_ORD -> CLOS_ORD
    PrepSO = ""
    If Not OpenByID(idv) Then log.WriteLine "  " & tag & " not found" : Exit Function
    log.WriteLine tag & " start: " & HState(0)
    If Hr(0,"STATU") = "C" And Hr(0,"VBELN") = "" Then
        r = HeaderAction(0, "CLOS_CANC", tag & "_clos_canc_log.png")
        If Not OpenByID(idv) Then Exit Function
        log.WriteLine "  after CLOS_CANC: " & HState(0)
    End If
    If Hr(0,"VBELN") = "" Then
        msg = DblHeader(0)
        log.WriteLine "  dblclick ATP: " & msg
        Shot tag & "_atp_dblclick.png"
        r = HeaderAction(0, "CREA_ORD", tag & "_crea_ord_log.png")
        log.WriteLine "  CREA_ORD: " & r
        If Not OpenByID(idv) Then Exit Function
    End If
    vb = Hr(0,"VBELN")
    log.WriteLine "  VBELN=" & vb & " " & HState(0)
    If vb <> "" And Hr(0,"STATU") <> "C" Then
        r = HeaderAction(0, "CLOS_ORD", tag & "_clos_ord_log.png")
        If Not OpenByID(idv) Then Exit Function
        log.WriteLine "  after CLOS_ORD: " & HState(0)
    End If
    PrepSO = Hr(0,"VBELN")
End Function

' ---------- 1722 ----------
vb = PrepSO(ID_1722, "C2_1722")
Result "PRE 1722 sales order + close", (vb <> ""), "VBELN, STATU C", "VBELN=" & vb & " " & HState(0)
If vb <> "" Then
    r = HeaderAction(0, "CRE_DELIVERY", "DIV_02_1722_cre_delivery_log.png")
    Shot "DIV_02_1722_after_button_inplace.png"
    Dim giInPlace : giInPlace = Hr(0,"GI_ICON")
    Dim vlInPlace : vlInPlace = Hr(0,"VBELN_VL")
    vl = ""
    If OpenByID(ID_1722) Then vl = Hr(0,"VBELN_VL") : r = r & " requery: " & HState(0)
    Shot "DIV_02_1722_requery.png"
    Result "DIV-02/DEL-02 1722 (parts KTGRM Y2) delivery + goods issue", (vl <> "" And Hr(0,"GI_ICON") <> ""), "delivery created AND GI posted (GI_ICON green)", r
    Result "ALV-05 1722 GI_ICON updated in place right after action", (giInPlace <> "" And vlInPlace <> ""), "GI_ICON green without re-query", "inplace VBELN_VL=[" & vlInPlace & "] GI_ICON=[" & giInPlace & "]"
    Result "ALV-06/ICON-04 1722 GI_ICON after fresh re-query", (Hr(0,"GI_ICON") <> ""), "green icon from VBUP-WBSTA", "GI_ICON=[" & Hr(0,"GI_ICON") & "]"
    If vl <> "" Then Se16n "VBUP", K_DOC, vl, "VBELN,POSNR,WBSTA,LFSTA,FKSTA", "ICON-03 VBUP " & vl, 12, "ICON_03_1722_vbup.png"
End If

' ---------- 2392 ----------
vb = PrepSO(ID_2392, "C2_2392")
Result "PRE 2392 sales order + close", (vb <> ""), "VBELN, STATU C", "VBELN=" & vb & " " & HState(0)

' ---------- 2827 (B, LGORT blank) ----------
vb = PrepSO(ID_2827, "C2_2827")
Note "PRE 2827 (LGORT blank) sales order + close", "VBELN=" & vb & " " & HState(0)

' ---------- DEL-03 : 2766 no sales order ----------
If OpenByID(ID_2766) Then
    log.WriteLine "2766: " & HState(0)
    r = HeaderAction(0, "CRE_DELIVERY", "DEL_03_2766_no_so_log.png")
    vl = "?"
    If OpenByID(ID_2766) Then vl = Hr(0,"VBELN_VL")
    Result "DEL-03 delivery on order without sales order (2766) blocked", (vl = ""), "blocked, error log", r
End If

' ---------- DEL-06 / REG-04 : mixed 2392 + 2827 + 1757(dup) ----------
If OpenByBSTKDList("2392,2827,1757") Then
    Dim n, rows, k
    n = HeaderRows()
    log.WriteLine "mixed query rows=" & n
    rows = ""
    For k = 0 To n - 1
        log.WriteLine "  H" & k & ": BSTKD=" & Hr(k,"BSTKD") & " " & HState(k)
        If Hr(k,"VBELN") <> "" Then
            If rows <> "" Then rows = rows & ","
            rows = rows & k
        End If
    Next
    SelHeaders rows
    session.findById(HG).pressToolbarButton "CRE_DELIVERY"
    WScript.Sleep 3000
    r = "selected=[" & rows & "] confirm=[" & WT(1) & "] " & YesAndLog("DEL_06_mixed_log.png")
    Shot "DEL_06_mixed_after.png"
    log.WriteLine "  DEL-06: " & r
    Dim s2392, s2827
    If OpenByID(ID_2392) Then s2392 = HState(0) : Shot "DEL_06_2392_after.png"
    If OpenByID(ID_2827) Then s2827 = HState(0) : Shot "DEL_06_2827_after.png"
    Result "DEL-06/REG-04 mixed delivery (2392 ok, 2827 LGORT blank, 1757 dup)", (InStr(s2392, "VBELN_VL=[]") = 0), "per-order success/failure, loop continues", "2392: " & s2392 & " | 2827: " & s2827 & " | log: " & r
    If InStr(s2392, "VBELN_VL=[]") = 0 Then
        Dim vl2392 : vl2392 = Mid(s2392, InStr(s2392, "VBELN_VL=[") + 10, 8)
        Se16n "VBUP", K_DOC, vl2392, "VBELN,POSNR,WBSTA,LFSTA,FKSTA", "2392 VBUP " & vl2392, 12, "DIV_01_2392_vbup.png"
    End If
Else
    log.WriteLine "mixed query failed: screen=" & session.Info.ScreenNumber & " sbar=" & Sbar()
End If

' ---------- ALV-08 : layout popups ----------
If OpenByID(ID_1722) Then
    Dim g : Set g = session.findById(HG)
    On Error Resume Next
    g.pressToolbarContextButton "&MB_VARIANT"
    WScript.Sleep 600
    g.selectContextMenuItem "&LOAD"
    WScript.Sleep 1500
    log.WriteLine "ALV-08 saved layouts popup=[" & WT(1) & "]"
    ShotW 1, "ALV_08_saved_layouts.png"
    CancelAll
    Set g = session.findById(HG)
    g.pressToolbarContextButton "&MB_VARIANT"
    WScript.Sleep 600
    g.selectContextMenuItem "&COL0"
    WScript.Sleep 1500
    log.WriteLine "ALV-08 change layout popup=[" & WT(1) & "]"
    ShotW 1, "ALV_08_change_layout.png"
    CancelAll
    Err.Clear
    On Error GoTo 0
    Note "ALV-08", "layout popups captured; check GI_ICON in displayed/hidden column lists in the shots"
End If
Finish
