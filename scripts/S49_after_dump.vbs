' S49 body : post-dump state of 1722 + VBUP 85857218 (DELETE-02/07 evidence). read only.
Const ID_1722 = "04a27fe5-6007-11f1-9ba0-000c298aeffe"
Dim vl
If OpenByID(ID_1722) Then
    log.WriteLine "1722 post-dump: " & HState(0)
    vl = Hr(0,"VBELN_VL")
    Shot "DELETE_02_1722_postdump_alv.png"
    Note "DELETE-02/07 1722 after GI-cancel-delete dump", "GI_ICON=[" & Hr(0,"GI_ICON") & "] VBELN_VL=[" & vl & "] (delivery still linked?)"
End If
Se16n "VBUP", K_DOC, "85857218", "VBELN,POSNR,WBSTA,LFSTA,FKSTA", "VBUP 85857218 post-dump", 12, "DELETE_02_vbup_postdump.png"
Se16n "LIKP", K_DOC, "85857218", "VBELN,ERDAT,WADAT_IST,WBSTK", "LIKP 85857218 post-dump", 3, ""
Result "DELETE-02 GI-completed delivery delete triggers ABAP dump (DEFECT)", False, "GI cancelled + delivery deleted", "ABAP dump SAPSQL_PARSER_TODO_WARNING in Z33SDR600: aggregate POSNR vs LV_LINES type incompatible (see DELETE_02_1722_ABAP_DUMP.png); delivery 85857218 state per VBUP above"
Finish
