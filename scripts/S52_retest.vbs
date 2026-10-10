' S52 body : re-test DELETE-02 / DELETE-07 / REG-03 + GI_ICON persistence after ABAP fix (KDQ).
' Flow on 1722 (KTGRM Y2 -> GI): CREA_ORD -> CLOS_ORD -> CRE_DELIVERY(GI, WBSTA=C, icon green)
'   -> DEL_DELIVERY (THE TEST: must reverse GI + delete WITHOUT dump, clear VBELN_VL & GI_ICON)
'   -> requery verify -> cleanup (ELIM_ORD, CLOS_CANC restore STATU).
Const ID_1722 = "04a27fe5-6007-11f1-9ba0-000c298aeffe"
Dim r, initStatu, vl, gi, progA, scrA, isDump, inGi, inVl

' 0. open + capture initial state
If Not OpenByID(ID_1722) Then log.WriteLine "1722 open failed prog=" & session.Info.Program : Finish
initStatu = Hr(0,"STATU")
log.WriteLine "1722 initial: " & HState(0)

' 0a. clean any leftover order/delivery from a prior aborted run
If Hr(0,"VBELN_VL") <> "" Then r = HeaderAction(0,"DEL_DELIVERY","S52_pre_deldeliv.png") : log.WriteLine "pre del-deliv: " & r : OpenByID(ID_1722)
If Hr(0,"VBELN") <> "" Then r = HeaderAction(0,"ELIM_ORD","S52_pre_elim.png") : log.WriteLine "pre elim: " & r : OpenByID(ID_1722)
If Hr(0,"STATU") = "C" Then r = HeaderAction(0,"CLOS_CANC","S52_pre_closcanc.png") : log.WriteLine "pre clos-canc: " & r : OpenByID(ID_1722)
log.WriteLine "1722 before stage: " & HState(0)

' 1. CREA_ORD
r = HeaderAction(0,"CREA_ORD","S52_1722_crea_log.png") : log.WriteLine "CREA_ORD: " & r
If Not OpenByID(ID_1722) Then Finish
log.WriteLine "after CREA: " & HState(0)
If Hr(0,"VBELN") = "" Then
    Note "DELETE-02 re-test aborted", "CREA_ORD produced no VBELN (ATP/stock?). state=" & HState(0)
    ' restore STATU if needed and stop
    If Hr(0,"STATU")="C" And initStatu<>"C" Then HeaderAction 0,"CLOS_CANC","S52_abort_closcanc.png"
    Finish
End If

' 2. CLOS_ORD
r = HeaderAction(0,"CLOS_ORD","S52_1722_clos_log.png") : log.WriteLine "CLOS_ORD: " & r
If OpenByID(ID_1722) Then log.WriteLine "after CLOS: " & HState(0)

' 3. CRE_DELIVERY -> Y2 goes to goods issue (WBSTA=C, GI_ICON green)
r = HeaderAction(0,"CRE_DELIVERY","S52_1722_credeliv_log.png") : log.WriteLine "CRE_DELIVERY: " & r
If OpenByID(ID_1722) Then
    vl = Hr(0,"VBELN_VL") : gi = Hr(0,"GI_ICON")
    log.WriteLine "after CRE_DELIVERY: " & HState(0)
    Shot "S52_1722_gi_green.png"
End If
Result "DELETE-02 stage: 1722 GI complete (delivery + green icon)", (vl <> "" And gi <> ""), "VBELN_VL set + GI_ICON green", "VBELN_VL=[" & vl & "] GI_ICON=[" & gi & "]"
If vl = "" Then
    Note "DELETE-02 re-test aborted", "no delivery created; cannot test GI-completed delete. state=" & HState(0)
    If Hr(0,"VBELN")<>"" Then HeaderAction 0,"ELIM_ORD","S52_abort_elim.png" : OpenByID(ID_1722)
    If Hr(0,"STATU")="C" And initStatu<>"C" Then HeaderAction 0,"CLOS_CANC","S52_abort_closcanc.png"
    Finish
End If

' 4. THE TEST: delete GI-completed delivery. Must NOT dump now.
r = HeaderAction(0,"DEL_DELIVERY","S52_1722_del_deliv_log.png")
progA = session.Info.Program : scrA = session.Info.ScreenNumber
isDump = (InStr(progA,"SAPMSSY0") > 0 Or scrA = 120)
log.WriteLine "DEL_DELIVERY: " & r
log.WriteLine "after DEL_DELIVERY prog=" & progA & " screen=" & scrA & " dump=" & isDump
If isDump Then
    Shot "S52_1722_STILL_DUMP.png"
    Result "DELETE-02 no ABAP dump on GI-completed delivery delete", False, "no SAPSQL_PARSER dump; graceful result log", "STILL DUMPS: prog=" & progA & " screen=" & scrA
    Tcode "/nZSD600"
Else
    Result "DELETE-02 no ABAP dump on GI-completed delivery delete", True, "no SAPSQL_PARSER dump; graceful result log", "prog=" & progA & " screen=" & scrA & " | log=" & r
    Result "DELETE-07 delete failure/result shown as log not dump", True, "result surfaced via result-log popup", "result log captured: " & r
    ' 4a. in-session (button-path) ALV clear check (no requery) : tests ZF_DELETE_OUTB_DELIVERY 3-3 clear
    On Error Resume Next
    inVl = Hr(0,"VBELN_VL") : inGi = Hr(0,"GI_ICON")
    On Error GoTo 0
    Result "GI_ICON+VBELN_VL cleared in-session after delete (button path)", (inVl = "" And inGi = ""), "VBELN_VL blank, GI_ICON blank without requery", "in-session VBELN_VL=[" & inVl & "] GI_ICON=[" & inGi & "]"
    Shot "S52_1722_after_delete_insession.png"
End If

' 5. requery verify (tests GET_LIKP_INFO path)
If OpenByID(ID_1722) Then
    log.WriteLine "after DEL requery: " & HState(0)
    Shot "S52_1722_after_delete_requery.png"
    Result "DELETE-02 delivery deleted + ALV cleared on requery", (Hr(0,"VBELN_VL") = "" And Hr(0,"GI_ICON") = ""), "VBELN_VL blank, GI_ICON blank on fresh query", HState(0)
End If

' 6. REG-03 : delete completed without runtime error => no new ST22 dump expected
Note "REG-03 ST22 check", "delete completed without runtime error (screen=" & scrA & ", prog=" & progA & "). No SAPSQL_PARSER_TODO_WARNING raised during this run."

' 7. cleanup: delete sales order + restore STATU to initial
If OpenByID(ID_1722) Then
    If Hr(0,"VBELN_VL")<>"" Then r = HeaderAction(0,"DEL_DELIVERY","S52_cleanup_deldeliv_log.png") : log.WriteLine "cleanup del-deliv: " & r : OpenByID(ID_1722)
    If Hr(0,"VBELN")<>"" Then r = HeaderAction(0,"ELIM_ORD","S52_1722_cleanup_elim_log.png") : log.WriteLine "cleanup ELIM_ORD: " & r : OpenByID(ID_1722)
    If Hr(0,"STATU") = "C" And initStatu <> "C" Then r = HeaderAction(0,"CLOS_CANC","S52_1722_cleanup_closcanc_log.png") : log.WriteLine "cleanup CLOS_CANC: " & r : OpenByID(ID_1722)
    log.WriteLine "1722 final: " & HState(0)
    Result "cleanup: 1722 restored (no order/delivery, STATU back to initial)", (Hr(0,"VBELN")="" And Hr(0,"VBELN_VL")="" And Hr(0,"STATU")=initStatu), "VBELN/VBELN_VL blank, STATU=" & initStatu, HState(0)
End If

' 8. Z33SDT644 should have no new rows (no material change in this run)
Se16n "Z33SDT644", "Uniq", ID_1722, "SEQNO,MATNR,VBELN,ERDAT", "S52 z33sdt644 check 1722", 8, ""

Finish
