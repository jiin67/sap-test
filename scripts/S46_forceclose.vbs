' S46 body : force-close 2827 via SE16N (Z33SDT640 STATU B->C), verify in ZSD600, then CRE_DELIVERY (LGORT blank -> expect error)
Const ID_2827 = "04bf872f-a20b-11f1-9bce-000c298aeffe"
Dim r, vl
' baseline
If OpenByID(ID_2827) Then log.WriteLine "2827 before: " & HState(0)
' force STATU = C
r = Se16nSetField("Z33SDT640", ID_2827, "STATU", "C", "S46_2827")
log.WriteLine "  SE16N: " & r
' verify
If OpenByID(ID_2827) Then log.WriteLine "2827 after SE16N: " & HState(0)
Result "PRE 2827 force STATU=C via SE16N (CLOS_ORD failed for LGORT-blank order)", (Hr(0,"STATU") = "C"), "STATU C", r & " | now " & HState(0)
' now delivery on LGORT-blank closed order -> expect BAPI error (R3-06 / DEL LGORT)
If Hr(0,"STATU") = "C" And Hr(0,"VBELN") <> "" Then
    r = HeaderAction(0, "CRE_DELIVERY", "S46_2827_cre_delivery_log.png")
    vl = "?"
    If OpenByID(ID_2827) Then vl = Hr(0,"VBELN_VL") : r = r & " after: " & HState(0)
    Result "R3-06/DEL 2827 delivery with blank storage location", (vl = ""), "delivery not created, BAPI/LGORT error in log", r
End If
Finish
