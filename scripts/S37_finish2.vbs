' S37_finish2.vbs : follow-ups. No permanent DB change.
'  DEL-04 proper (order WITH sales order, WITHOUT delivery: 1422 / VBELN 4293032) -> Yes -> message + log popup shot (UI-06 evidence)
'  MARC check for PNGB3500035L001 and 20010472A (SE16N, read only) -> substantiate plant-extension finding
'  HIST-10 detail: duplicate header rows on material query, and same BSTKD queried without material
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i, nPass, nFail
nPass = 0 : nFail = 0
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S37_finish2.txt", True, True)
If connection.Children.Count > 1 Then
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
Else
    n0 = connection.Children.Count
    base.createSession
    For i = 1 To 40
        WScript.Sleep 500
        If connection.Children.Count > n0 Then Exit For
    Next
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    WScript.Sleep 1500
End If
log.WriteLine "System=" & session.Info.SystemName & " Client=" & session.Info.Client & " User=" & session.Info.User & " RunAt=" & Now
Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"
Dim K_YES, K_DELIV, K_MATNRDESC, RU_OT, BSTKD_1422
K_YES = ChrW(&HC608)
K_DELIV = ChrW(&HB0A9)&ChrW(&HD488)
K_MATNRDESC = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HBC88)&ChrW(&HD638)   ' jaejae-beonho
RU_OT = ChrW(&H43E)&ChrW(&H442)
BSTKD_1422 = "1422 " & RU_OT & " 12.05.2026"

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Function Sbar() : Sbar = session.findById("wnd[0]/sbar").Text : End Function
Function SbarT() : SbarT = session.findById("wnd[0]/sbar").MessageType : End Function
Function WT(n)
    On Error Resume Next
    WT = "" : WT = session.findById("wnd[" & n & "]").Text
    Err.Clear
    On Error GoTo 0
End Function
Function HasW(n)
    On Error Resume Next
    HasW = False
    Dim w : Set w = session.findById("wnd[" & n & "]")
    If Err.Number = 0 Then HasW = True
    Err.Clear
    On Error GoTo 0
End Function
Function W2Msg()
    On Error Resume Next
    W2Msg = "" : W2Msg = session.findById("wnd[2]/usr/txtMESSTXT1").Text
    Err.Clear
    On Error GoTo 0
End Function
Sub CancelAll()
    Dim g, n, usr, c, comp, did
    On Error Resume Next
    For g = 1 To 8
        If Not HasW(1) Then Exit Sub
        For n = 3 To 1 Step -1
            If HasW(n) Then
                did = False
                Set usr = session.findById("wnd[" & n & "]/usr")
                If Err.Number = 0 Then
                    For c = 0 To usr.Children.Count - 1
                        Set comp = usr.Children.ElementAt(c)
                        If comp.Type = "GuiButton" Then
                            If InStr(comp.Text & comp.Tooltip, ChrW(&HC544)&ChrW(&HB2C8)&ChrW(&HC624))>0 Or InStr(comp.Text & comp.Tooltip, ChrW(&HCDE8)&ChrW(&HC18C))>0 Then comp.press : did = True : Exit For
                        End If
                    Next
                End If
                Err.Clear
                If Not did Then session.findById("wnd[" & n & "]").sendVKey 0 : Err.Clear
                If HasW(n) Then session.findById("wnd[" & n & "]/tbar[0]/btn[12]").press : Err.Clear
                If HasW(n) Then session.findById("wnd[" & n & "]").close : Err.Clear
                WScript.Sleep 500
            End If
        Next
    Next
    On Error GoTo 0
End Sub
Sub Result(tc, ok, expected, actual)
    If ok Then nPass = nPass+1 : log.WriteLine "[PASS] " & tc Else nFail = nFail+1 : log.WriteLine "[FAIL] " & tc
    log.WriteLine "       expect: " & expected
    log.WriteLine "       actual: " & actual
End Sub
Function OpenQuery(bstkd, matnr)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = bstkd
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = matnr
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 4000
    OpenQuery = (session.Info.ScreenNumber = 100)
End Function
' dump rows of a result-log popup grid, and screenshot popup
Function LogPopupRows(shot)
    Dim s, g, r, cols, cc, ln
    s = ""
    On Error Resume Next
    If HasW(1) Then
        session.findById("wnd[1]").HardCopy shot, 2
        Set g = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
        If Err.Number = 0 Then
            Set cols = g.ColumnOrder
            s = "logRows=" & g.RowCount
            For r = 0 To g.RowCount - 1
                If r > 4 Then Exit For
                ln = " | row" & r & ":"
                For cc = 0 To cols.Count - 1
                    If cols(cc)="BSTKD" Or cols(cc)="VBELN_VA" Or cols(cc)="VBELN_VL" Or cols(cc)="RESULT" Or cols(cc)="RESULT_MESSAGE" Or cols(cc)="MESSAGE" Then ln = ln & " " & cols(cc) & "=[" & g.GetCellValue(r, cols(cc)) & "]"
                Next
                s = s & ln
            Next
        End If
    End If
    Err.Clear
    On Error GoTo 0
    LogPopupRows = s
End Function
Sub Se16nByMatnr(table, matnr, label)
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "500"
    Dim rr, desc, setC : setC = False
    For rr = 0 To 25
        desc = "" : desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, K_MATNRDESC) > 0 Then session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = matnr : setC = True : Exit For
    Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2500
    Dim g, r2, ln
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then log.WriteLine "  [" & label & "] " & table & " MATNR=" & matnr & " condSet=" & setC & " -> no ALV, sbar=" & Sbar() : Err.Clear : On Error GoTo 0 : Exit Sub
    log.WriteLine "  [" & label & "] " & table & " MATNR=" & matnr & " condSet=" & setC & " rows=" & g.RowCount
    For r2 = 0 To g.RowCount - 1
        If r2 > 15 Then Exit For
        ln = "   #" & r2 & ": MATNR=[" & g.GetCellValue(r2,"MATNR") & "]"
        If table = "MARC" Then ln = ln & " WERKS=[" & g.GetCellValue(r2,"WERKS") & "] MMSTA=[" & g.GetCellValue(r2,"MMSTA") & "] LVORM=[" & g.GetCellValue(r2,"LVORM") & "] MTVFP=[" & g.GetCellValue(r2,"MTVFP") & "]"
        If table = "MARA" Then ln = ln & " MTART=[" & g.GetCellValue(r2,"MTART") & "] MEINS=[" & g.GetCellValue(r2,"MEINS") & "] LVORM=[" & g.GetCellValue(r2,"LVORM") & "]"
        log.WriteLine ln
    Next
    On Error GoTo 0
End Sub

' ================= DEL-04 proper: sales order exists, no delivery =================
Sub DEL_04b()
    Dim tc : tc = "DEL-04 delete-delivery on order WITH sales order, WITHOUT delivery (1422/4293032)"
    If Not OpenQuery(BSTKD_1422, "") Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber & " " & Sbar() : Exit Sub
    Dim gh : Set gh = session.findById(HG)
    Dim vb : vb = gh.GetCellValue(0,"VBELN")
    Dim vl : vl = gh.GetCellValue(0,"VBELN_VL")
    log.WriteLine "  1422: headerRows=" & gh.RowCount & " VBELN=[" & vb & "] VBELN_VL=[" & vl & "] STATU=" & gh.GetCellValue(0,"STATU")
    If vb = "" Or vl <> "" Then Result tc, False, "VBELN set and VBELN_VL empty", "VBELN=[" & vb & "] VL=[" & vl & "]" : Exit Sub
    gh.setCurrentCell 0,"BSTKD" : gh.selectedRows = "0"
    gh.pressToolbarButton "DEL_DELIVERY"
    WScript.Sleep 2000
    Dim pt : pt = WT(1)
    Dim usr, c, comp, py : py = False
    On Error Resume Next
    Set usr = session.findById("wnd[1]/usr")
    If Err.Number = 0 Then
        For c = 0 To usr.Children.Count - 1
            Set comp = usr.Children.ElementAt(c)
            If comp.Type = "GuiButton" Then
                If InStr(comp.Text & comp.Tooltip, K_YES) > 0 Then comp.press : py = True : Exit For
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 2500
    Dim msg : msg = "confirm=[" & pt & "] yes=" & py & " sbar[" & SbarT() & ":" & Sbar() & "] w1=[" & WT(1) & "] w2=[" & W2Msg() & "] " & LogPopupRows("C:\sap-test\results\shots\DEL_04_log_popup.png")
    CancelAll
    Dim vb2, vl2 : vb2 = "?" : vl2 = "?"
    If OpenQuery(BSTKD_1422, "") Then vb2 = session.findById(HG).GetCellValue(0,"VBELN") : vl2 = session.findById(HG).GetCellValue(0,"VBELN_VL")
    Result tc, (py And InStr(msg, K_DELIV) > 0 And vb2 = vb And vl2 = ""), "'no delivery to delete' message; VBELN " & vb & " untouched", msg & " after VBELN=[" & vb2 & "] VL=[" & vl2 & "]"
End Sub

' ================= MARC / MARA check =================
Sub MarcCheck()
    log.WriteLine "=== MARA/MARC: PNGB3500035L001 vs 20010472A (material of order 1422, which has a sales order) ==="
    Se16nByMatnr "MARA", "PNGB3500035L001", "mara-pngb"
    Se16nByMatnr "MARC", "PNGB3500035L001", "marc-pngb"
    Se16nByMatnr "MARC", "20010472A", "marc-1422mat"
    Tcode "/n"
End Sub

' ================= HIST-10 detail =================
Sub HIST_10b()
    Dim tc : tc = "HIST-10 duplicate header analysis (material GXXX002868)"
    If Not OpenQuery("", "GXXX002868") Then Result tc, False, "result", "screen=" & session.Info.ScreenNumber : Exit Sub
    Dim gh : Set gh = session.findById(HG)
    Dim d : Set d = CreateObject("Scripting.Dictionary")
    Dim r, key, firstDupKey, firstDupBstkd
    firstDupKey = "" : firstDupBstkd = ""
    For r = 0 To gh.RowCount - 1
        key = gh.GetCellValue(r,"ID") & "|" & gh.GetCellValue(r,"BSTKD") & "|" & gh.GetCellValue(r,"VERSION")
        If d.Exists(key) Then
            d(key) = d(key) & "," & r
            If firstDupKey = "" Then firstDupKey = key : firstDupBstkd = gh.GetCellValue(r,"BSTKD")
        Else
            d.Add key, CStr(r)
        End If
    Next
    Dim k, dupKeys : dupKeys = 0
    For Each k In d.Keys
        If InStr(d(k), ",") > 0 Then
            dupKeys = dupKeys + 1
            If dupKeys <= 3 Then
                Dim parts, p, rr
                parts = Split(d(k), ",")
                log.WriteLine "  dupKey " & k & " rows=" & d(k)
                For p = 0 To UBound(parts)
                    rr = CInt(parts(p))
                    log.WriteLine "     row" & rr & ": STATU=" & gh.GetCellValue(rr,"STATU") & " VBELN=[" & gh.GetCellValue(rr,"VBELN") & "] VL=[" & gh.GetCellValue(rr,"VBELN_VL") & "] CREDAT=" & gh.GetCellValue(rr,"CREDAT") & " NETPR_SUM=" & gh.GetCellValue(rr,"NETPR_SUM") & " AEDAT=" & gh.GetCellValue(rr,"AEDAT") & " LOEVM=[" & gh.GetCellValue(rr,"LOEVM") & "]"
                Next
            End If
        End If
    Next
    log.WriteLine "  material query: headerRows=" & gh.RowCount & " uniqueKeys=" & d.Count & " keysWithDuplicates=" & dupKeys
    ' item count of the first duplicated header: how many items carry the material?
    Dim itemsWithMat : itemsWithMat = -1
    Dim nH : nH = 0
    If firstDupBstkd <> "" Then
        If OpenQuery(firstDupBstkd, "") Then
            nH = session.findById(HG).RowCount
            Dim gh2 : Set gh2 = session.findById(HG)
            Dim h
            For h = 0 To gh2.RowCount - 1
                If gh2.GetCellValue(h,"ID") & "|" & gh2.GetCellValue(h,"BSTKD") & "|" & gh2.GetCellValue(h,"VERSION") = firstDupKey Then
                    gh2.setCurrentCell h,"BSTKD" : gh2.selectedRows = CStr(h)
                    On Error Resume Next
                    gh2.doubleClickCurrentCell
                    Err.Clear
                    On Error GoTo 0
                    WScript.Sleep 1800
                    CancelAll
                    Dim gi : Set gi = session.findById(IG)
                    Dim it : itemsWithMat = 0
                    For it = 0 To gi.RowCount - 1
                        If gi.GetCellValue(it,"MATNR") = "GXXX002868" Then itemsWithMat = itemsWithMat + 1
                    Next
                    log.WriteLine "  BSTKD-only query for " & firstDupBstkd & ": headerRows=" & nH & "; key " & firstDupKey & " items=" & gi.RowCount & " itemsWithMaterial=" & itemsWithMat
                    Exit For
                End If
            Next
        End If
    End If
    Result tc, (dupKeys = 0), "each ID|BSTKD|VERSION once", "headerRows=" & gh.RowCount & " unique=" & d.Count & " dupKeys=" & dupKeys & " (first dup " & firstDupBstkd & ": itemsWithMaterial=" & itemsWithMat & ", BSTKD-only headerRows=" & nH & ")"
    Tcode "/n"
End Sub

On Error Resume Next
DEL_04b : If Err.Number <> 0 Then Result "DEL-04", False, "-", "runtime " & Err.Description : Err.Clear
CancelAll
MarcCheck : If Err.Number <> 0 Then log.WriteLine "  MarcCheck runtime " & Err.Description : Err.Clear
CancelAll
HIST_10b : If Err.Number <> 0 Then Result "HIST-10", False, "-", "runtime " & Err.Description : Err.Clear
CancelAll
On Error GoTo 0
Tcode "/n"
log.WriteLine "SUMMARY: PASS=" & nPass & " FAIL=" & nFail
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "S37 done PASS=" & nPass & " FAIL=" & nFail
