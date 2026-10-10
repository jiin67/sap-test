' S51 body : restore original data. delete Z33SDT644 test rows (1757, 2318); restore STATU (2318->C, 2827->B). verify.
Const ID_1757 = "03f32fd8-6329-11f1-9ba3-000c298aeffe"
Const ID_2318 = "0407e0e8-8434-11f1-9bbb-000c298aeffe"
Const ID_2827 = "04bf872f-a20b-11f1-9bce-000c298aeffe"
Dim r
r = Se16nDeleteByID("Z33SDT644", ID_1757, "S51_644_1757") : log.WriteLine "del644 1757: " & r
r = Se16nDeleteByID("Z33SDT644", ID_2318, "S51_644_2318") : log.WriteLine "del644 2318: " & r
r = Se16nSetField("Z33SDT640", ID_2318, "STATU", "C", "S51_statu_2318") : log.WriteLine "statu 2318: " & r
r = Se16nSetField("Z33SDT640", ID_2827, "STATU", "B", "S51_statu_2827") : log.WriteLine "statu 2827: " & r
' verify Z33SDT644 empty + STATU + item materials restored
Se16n "Z33SDT644", "Uniq", ID_1757, "SEQNO,MATNR,VBELN", "verify644 1757", 5, ""
Se16n "Z33SDT644", "Uniq", ID_2318, "SEQNO,MATNR,VBELN", "verify644 2318", 5, ""
Dim msg
If OpenByID(ID_1757) Then msg = DblHeader(0) : log.WriteLine "1757 final: " & HState(0) & " item0=" & It(0,"MATNR")
If OpenByID(ID_2318) Then msg = DblHeader(0) : log.WriteLine "2318 final: " & HState(0) & " item11=" & It(10,"MATNR")
If OpenByID(ID_2827) Then log.WriteLine "2827 final: " & HState(0)
Note "cleanup", "all created sales orders/deliveries removed; Z33SDT644 test rows deleted; STATU restored (2318->C, 2827->B)"
Finish
