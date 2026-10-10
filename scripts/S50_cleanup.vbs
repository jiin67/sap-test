' S50 body : delete remaining deliveries (now WBSTA A) + ELIM_ORD sweep of created sales orders + Z33SDT644 check. cleanup.
Dim IDS, TAGS, k, r, idv, tag
IDS = Array("03f32fd8-6329-11f1-9ba3-000c298aeffe", _
            "0407e0e8-8434-11f1-9bbb-000c298aeffe", _
            "04a27fe5-6007-11f1-9ba0-000c298aeffe", _
            "04a5f20c-8a85-11f1-9bc0-000c298aeffe", _
            "04bf872f-a20b-11f1-9bce-000c298aeffe")
TAGS = Array("1757","2318","1722","2392","2827")

' step 1: delete any remaining deliveries (WBSTA should be A now)
For k = 0 To UBound(IDS)
    idv = IDS(k) : tag = TAGS(k)
    If OpenByID(idv) Then
        If Hr(0,"VBELN_VL") <> "" Then
            r = HeaderAction(0, "DEL_DELIVERY", "S50_" & tag & "_deldeliv_log.png")
            If OpenByID(idv) Then log.WriteLine tag & " del-delivery: " & HState(0) & " | " & r
            If tag = "1722" Then Result "DELETE-02-retry 1722 delivery (WBSTA A after GI cancel) deletes cleanly", (Hr(0,"VBELN_VL") = ""), "delivery deleted after GI already cancelled", r
        End If
    End If
Next

' step 2: ELIM_ORD sweep (delete sales orders)
For k = 0 To UBound(IDS)
    idv = IDS(k) : tag = TAGS(k)
    If OpenByID(idv) Then
        If Hr(0,"VBELN") <> "" Then
            r = HeaderAction(0, "ELIM_ORD", "S50_" & tag & "_elim_log.png")
            If OpenByID(idv) Then log.WriteLine tag & " ELIM_ORD: " & HState(0) & " | " & r
        Else
            log.WriteLine tag & " no VBELN: " & HState(0)
        End If
    End If
Next

' step 3: Z33SDT644 leftover check + delete rows created today (SEQNO change history)
For k = 0 To UBound(IDS)
    idv = IDS(k) : tag = TAGS(k)
    Se16n "Z33SDT644", "Uniq", idv, "SEQNO,MATNR_OLD,MATNR,VBELN,ERDAT", "cleanup644 " & tag, 12, ""
Next

' final state
log.WriteLine "=== final states ==="
For k = 0 To UBound(IDS)
    If OpenByID(IDS(k)) Then log.WriteLine TAGS(k) & ": " & HState(0)
Next
Finish
