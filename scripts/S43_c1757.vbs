' S43 body : 1757 (B, SO 4293338, KTGRM Y1) -> DEL-05 (unclosed blocked) -> CLOS_ORD -> CRE_DELIVERY (DIV-01/DEL-01/DEL-07)
'            -> DEL-04 duplicate -> VBUP (ICON-01) -> ALV-02 ; ALV-01 column evidence
Const ID_1757 = "03f32fd8-6329-11f1-9ba3-000c298aeffe"
Dim r, msg, vl, vb
If Not OpenByID(ID_1757) Then log.WriteLine "ABORT 1757" : Finish : WScript.Quit 10
log.WriteLine "1757 start: " & HState(0)
vb = Hr(0,"VBELN")
If vb = "" Then log.WriteLine "ABORT no VBELN" : Finish : WScript.Quit 10

' ALV-01 : GI_ICON column present with title
Dim gh, cols, c, hasGI, giTitle
Set gh = session.findById(HG) : Set cols = gh.ColumnOrder : hasGI = False
For c = 0 To cols.Count - 1
    If cols(c) = "GI_ICON" Then hasGI = True : giTitle = gh.GetColumnTitles(cols(c))(0)
Next
Result "ALV-01 GI_ICON column present, title", (hasGI And giTitle = ChrW(&HCD9C)&ChrW(&HACE0)&ChrW(&HC804)&ChrW(&HAE30)), "column GI_ICON titled chulgo-jeongi", "present=" & hasGI & " title=[" & giTitle & "]"

' DEL-05 : STATU B (not closed) -> delivery creation blocked
If Hr(0,"STATU") <> "C" Then
    r = HeaderAction(0, "CRE_DELIVERY", "DEL_05_1757_unclosed_log.png")
    vl = "?"
    If OpenByID(ID_1757) Then vl = Hr(0,"VBELN_VL") : r = r & " after: " & HState(0)
    Result "DEL-05 delivery on unclosed order (STATU B) blocked", (vl = ""), "blocked, 'not closed' message", r
End If

' CLOS_ORD -> C
r = HeaderAction(0, "CLOS_ORD", "C01_1757_clos_ord_log.png")
If OpenByID(ID_1757) Then r = r & " after: " & HState(0)
Result "PRE 1757 CLOS_ORD -> STATU C", (Hr(0,"STATU") = "C"), "STATU C", r

' DEL-01 / DIV-01 : delivery only (KTGRM Y1)
r = HeaderAction(0, "CRE_DELIVERY", "DEL_01_1757_cre_delivery_log.png")
Shot "DEL_01_1757_after_button.png"
vl = ""
If OpenByID(ID_1757) Then vl = Hr(0,"VBELN_VL") : r = r & " requery: " & HState(0)
Shot "DEL_01_1757_requery.png"
Result "DEL-01/DIV-01 1757 (all KTGRM Y1) delivery created, GI_ICON blank", (vl <> "" And Hr(0,"GI_ICON") = ""), "VBELN_VL set, GI_ICON blank, one log row per document", r
Note "DEL-07 log rows for 1757 delivery", "see logRows above (2 items in order)"

' VBUP of the delivery -> ICON-01 / ALV-02
If vl <> "" Then
    Se16n "VBUP", K_DOC, vl, "VBELN,POSNR,WBSTA,LFSTA,FKSTA,KOSTA", "VBUP " & vl, 12, "ICON_01_1757_vbup.png"
    Se16n "LIPS", K_DOC, vl, "VBELN,POSNR,VGBEL,VGPOS,MATNR,LFIMG,LGORT,WERKS", "LIPS " & vl, 12, "DEL_01_1757_lips.png"
    ' DEL-04 : duplicate delivery blocked
    If OpenByID(ID_1757) Then
        r = HeaderAction(0, "CRE_DELIVERY", "DEL_04_1757_duplicate_log.png")
        Dim vl2 : vl2 = "?"
        If OpenByID(ID_1757) Then vl2 = Hr(0,"VBELN_VL") : r = r & " after: " & HState(0)
        Result "DEL-04 duplicate delivery blocked (existing " & vl & ")", (vl2 = vl), "no new delivery, existing number in log", r
        Result "ALV-02/ICON-01 1757 delivery WBSTA A -> GI_ICON blank", (Hr(0,"GI_ICON") = ""), "GI_ICON blank", "GI_ICON=[" & Hr(0,"GI_ICON") & "] (WBSTA per VBUP above)"
        Shot "ALV_02_1757_gi_blank.png"
    End If
End If
Finish
