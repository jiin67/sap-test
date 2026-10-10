' S47 body : delivery deletion + GI cancel lifecycle, history verify, header de-dup
'   DELETE-01/ALV-07 : 1757 delivery 85857217 (WBSTA A) -> DEL_DELIVERY -> deleted, VBELN_VL & GI blank
'   DELETE-02       : 1722 delivery 85857218 (WBSTA C, GI'd) -> DEL_DELIVERY -> GI cancel + delete, VBUP back to A
'   DELETE-05       : 1757 (now no delivery) -> DEL_DELIVERY -> 'no delivery' message
'   2827 delivery 85857219 -> delete (cleanup)
'   HIST-02/03 2318 : SO 4293339 exists (item11 changed) -> Z33SDT644 + re-query
'   REG-01          : header de-dup by material query
Const ID_1757 = "03f32fd8-6329-11f1-9ba3-000c298aeffe"
Const ID_1722 = "04a27fe5-6007-11f1-9ba0-000c298aeffe"
Const ID_2827 = "04bf872f-a20b-11f1-9bce-000c298aeffe"
Const ID_2318 = "0407e0e8-8434-11f1-9bbb-000c298aeffe"
Dim r, vl, gi

' DELETE-01 : 1757 WBSTA A -> delete
If OpenByID(ID_1757) Then
    vl = Hr(0,"VBELN_VL")
    log.WriteLine "1757 before delete: " & HState(0)
    If vl <> "" Then
        r = HeaderAction(0, "DEL_DELIVERY", "DELETE_01_1757_log.png")
        Shot "DELETE_01_1757_after_inplace.png"
        Dim vlIn : vlIn = Hr(0,"VBELN_VL")
        If OpenByID(ID_1757) Then r = r & " requery: " & HState(0)
        Shot "DELETE_01_1757_requery.png"
        Result "DELETE-01/ALV-07 1757 unshipped delivery (" & vl & ") deleted, VBELN_VL & GI cleared", (Hr(0,"VBELN_VL") = "" And Hr(0,"GI_ICON") = ""), "VBELN_VL blank, GI blank", "inplaceVL=[" & vlIn & "] " & r
        ' DELETE-05 : now no delivery -> delete again
        r = HeaderAction(0, "DEL_DELIVERY", "DELETE_05_1757_none_log.png")
        Result "DELETE-05 1757 delete with no delivery -> 'none' message", (InStr(r, K_DOC) > 0 Or InStr(r, ChrW(&HC5C6)&ChrW(&HC2B5)) > 0 Or InStr(r, "RESULT=[E]") > 0), "no-delivery message", r
    Else
        Note "DELETE-01 1757", "no delivery present: " & HState(0)
    End If
End If

' DELETE-02 : 1722 WBSTA C (GI'd) -> GI cancel + delete
If OpenByID(ID_1722) Then
    vl = Hr(0,"VBELN_VL") : gi = Hr(0,"GI_ICON")
    log.WriteLine "1722 before delete: " & HState(0)
    If vl <> "" Then
        r = HeaderAction(0, "DEL_DELIVERY", "DELETE_02_1722_log.png")
        Shot "DELETE_02_1722_after_inplace.png"
        Dim vlIn2, giIn2 : vlIn2 = Hr(0,"VBELN_VL") : giIn2 = Hr(0,"GI_ICON")
        If OpenByID(ID_1722) Then r = r & " requery: " & HState(0)
        Shot "DELETE_02_1722_requery.png"
        Result "DELETE-02 1722 GI-completed delivery (" & vl & ", GI " & gi & ") -> GI cancel + delete", (Hr(0,"VBELN_VL") = "" And Hr(0,"GI_ICON") = ""), "GI cancelled, delivery deleted, VBELN_VL & GI blank", "beforeGI=[" & gi & "] inplaceVL=[" & vlIn2 & "] inplaceGI=[" & giIn2 & "] " & r
    End If
End If

' 2827 delivery cleanup
If OpenByID(ID_2827) Then
    vl = Hr(0,"VBELN_VL")
    If vl <> "" Then
        r = HeaderAction(0, "DEL_DELIVERY", "S47_2827_del_log.png")
        If OpenByID(ID_2827) Then r = r & " after: " & HState(0)
        Note "cleanup 2827 delivery " & vl & " deleted", r
    End If
End If

' HIST-02/03 : 2318 (SO 4293339, item11 changed earlier)
If OpenByID(ID_2318) Then
    log.WriteLine "2318: " & HState(0)
    Se16n "Z33SDT644", "Uniq", ID_2318, "SEQNO,MATNR_OLD,MATNR,VBELN,ERNAM,ERDAT,ERZET", "HIST-02 Z33SDT644 2318", 12, "HIST_02_2318_z33sdt644.png"
    Dim msg : msg = DblHeader(0)
    log.WriteLine "  2318 items: item11(SEQNO11)=" & It(10,"MATNR")
    Shot "HIST_03_2318_requery_items.png"
    Result "HIST-02/03 2318 changed material persisted + shown after re-query", (Hr(0,"VBELN") <> "" And It(10,"MATNR") = "PNGB3500013L008"), "Z33SDT644 has row, item11 shows changed material", "VBELN=" & Hr(0,"VBELN") & " item11=" & It(10,"MATNR")
End If

' REG-01 : header de-dup by material
Tcode "/nZSD600"
If WaitSel() Then
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = "PNGB3500013L008"
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 4000
    If session.Info.ScreenNumber = 100 Then
        Dim n, d, dup, rr, key
        n = HeaderRows()
        Set d = CreateObject("Scripting.Dictionary")
        dup = 0
        Dim vis : vis = session.findById(HG).VisibleRowCount
        Dim pg
        For pg = 0 To n - 1 Step vis
            session.findById(HG).FirstVisibleRow = pg
            WScript.Sleep 200
            For rr = pg To pg + vis - 1
                If rr >= n Then Exit For
                key = Hr(rr,"ID") & "|" & Hr(rr,"BSTKD") & "|" & Hr(rr,"VERSION")
                If Hr(rr,"ID") <> "" Then
                    If d.Exists(key) Then dup = dup + 1 Else d.Add key, 1
                End If
            Next
        Next
        Result "REG-01 header de-dup (material PNGB3500013L008 query)", (dup = 0 And n > 0), "each ID|BSTKD|VERSION once", "rows=" & n & " uniqueKeys=" & d.Count & " duplicates=" & dup
    End If
End If
Finish
