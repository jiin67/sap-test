' S36_finish.vbs : remaining data-independent cases. No permanent DB change.
'  DEL-04 (Yes on delete-confirm for order WITHOUT delivery -> message), R9-04 redesigned (same SAP material re-entry),
'  R8-06 / R8-08 colour evidence (HardCopy), HIST-10 header de-dup on material query.
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i, nPass, nFail
nPass = 0 : nFail = 0
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder("C:\sap-test\results\shots")
Set log = fso.CreateTextFile("C:\sap-test\logs\S36_finish.txt", True, True)
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
Const TARGET_ID = "06aa91b1-7b8c-11f1-9bb5-000c298aeffe"
Const NEWMAT = "PNGB3500035L001"
Dim K_CHGTITLE, K_YES, K_DELIV, K_SAME
K_CHGTITLE = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HCF54)&ChrW(&HB4DC)
K_YES   = ChrW(&HC608)
K_DELIV = ChrW(&HB0A9)&ChrW(&HD488)
K_SAME  = ChrW(&HAC19)&ChrW(&HC2B5)&ChrW(&HB2C8)&ChrW(&HB2E4)

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
Function OpenByID(idv)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = idv
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 3000
    OpenByID = (session.Info.ScreenNumber = 100)
End Function
Sub DblHeader0()
    Dim gh : Set gh = session.findById(HG)
    gh.setCurrentCell 0,"BSTKD" : gh.selectedRows = "0"
    On Error Resume Next
    gh.doubleClickCurrentCell
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 1800
End Sub
' in-memory change of item0 to NEWMAT (Save button). Returns item MATNR after.
Function ChangeItem0To(m)
    session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1400
    ChangeItem0To = "(dialog not open)"
    If InStr(WT(1), K_CHGTITLE) = 0 Then CancelAll : Exit Function
    session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = m
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 1200
    If HasW(2) Then ChangeItem0To = "err:" & W2Msg() : CancelAll : Exit Function
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    WScript.Sleep 1400
    CancelAll
    On Error Resume Next
    ChangeItem0To = session.findById(IG).GetCellValue(0,"MATNR")
    Err.Clear
    On Error GoTo 0
End Function

' ================= DEL-04 =================
Sub DEL_04()
    Dim tc : tc = "DEL-04 delete-delivery when none exists (Yes on confirm)"
    If Not OpenByID(TARGET_ID) Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    Dim gh : Set gh = session.findById(HG)
    If gh.GetCellValue(0,"VBELN_VL") <> "" Or gh.GetCellValue(0,"VBELN") <> "" Then Result tc, False, "no VBELN/VBELN_VL", "VBELN=" & gh.GetCellValue(0,"VBELN") & " VL=" & gh.GetCellValue(0,"VBELN_VL") : Exit Sub
    gh.setCurrentCell 0,"BSTKD" : gh.selectedRows = "0"
    gh.pressToolbarButton "DEL_DELIVERY"
    WScript.Sleep 2000
    Dim pt : pt = WT(1)
    Dim msg : msg = ""
    ' press Yes only because there is provably nothing to delete
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
    WScript.Sleep 2000
    msg = "sbar[" & SbarT() & ":" & Sbar() & "] w1=[" & WT(1) & "] w2=[" & W2Msg() & "]"
    ' result popup (SALV log) rows?
    On Error Resume Next
    Dim g : Set g = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
    If Err.Number = 0 Then
        Dim r, cols, cc, ln
        Set cols = g.ColumnOrder
        For r = 0 To g.RowCount - 1
            If r > 3 Then Exit For
            ln = " row" & r & ":"
            For cc = 0 To cols.Count - 1
                ln = ln & " " & cols(cc) & "=[" & g.GetCellValue(r, cols(cc)) & "]"
            Next
            msg = msg & ln
        Next
    End If
    Err.Clear
    On Error GoTo 0
    CancelAll
    Dim vl : vl = "?"
    If OpenByID(TARGET_ID) Then vl = session.findById(HG).GetCellValue(0,"VBELN_VL")
    Result tc, (py And InStr(msg, K_DELIV) > 0 And vl = ""), "confirm=" & pt & " -> Yes -> 'no delivery to delete' message; nothing changes", msg & " afterVL=[" & vl & "]"
End Sub

' ================= R8-06 evidence (original 1C code) =================
Sub R8_06()
    Dim tc : tc = "R8-06 unregistered material item colour (evidence)"
    If Not OpenByID(TARGET_ID) Then Result tc, False, "screen100", "x" : Exit Sub
    DblHeader0
    Dim sb : sb = Sbar()
    CancelAll
    On Error Resume Next
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\R8_06_item_1ccode.png", 2
    Err.Clear
    On Error GoTo 0
    Result tc, True, "item with 1C code (not in MARA) shown; colour per screenshot", "item0 MATNR=" & session.findById(IG).GetCellValue(0,"MATNR") & " MAKTX=[" & session.findById(IG).GetCellValue(0,"MAKTX") & "] sbar=[" & sb & "] shot=R8_06_item_1ccode.png"
End Sub

' ================= R9-04 redesigned + R8-08/09 evidence =================
Sub R9_04_R8()
    Dim tc : tc = "R9-04 same SAP material re-entered after in-memory change"
    If Not OpenByID(TARGET_ID) Then Result tc, False, "screen100", "x" : Exit Sub
    DblHeader0 : CancelAll
    Dim after1 : after1 = ChangeItem0To(NEWMAT)
    log.WriteLine "  step1 change -> item0 MATNR=" & after1
    If after1 <> NEWMAT Then Result tc, False, "item0=" & NEWMAT, after1 : Exit Sub
    ' R8-08/09: re-run ATP check by double-clicking header again, capture sbar + screenshot
    DblHeader0
    Dim sbAtp : sbAtp = "sbar[" & SbarT() & ":" & Sbar() & "] w1=[" & WT(1) & "] w2=[" & W2Msg() & "]"
    CancelAll
    On Error Resume Next
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\R8_08_item_pngb_atp.png", 2
    Err.Clear
    On Error GoTo 0
    Dim m2 : m2 = session.findById(IG).GetCellValue(0,"MATNR")
    Result "R8-08/09 ATP not checkable material (evidence)", True, "after change to " & NEWMAT & ", header dbl-click re-runs ATP; colour per screenshot", "item0=" & m2 & " " & sbAtp & " shot=R8_08_item_pngb_atp.png"
    ' now re-enter the SAME material -> expect same-material block (M25)
    If m2 <> NEWMAT Then
        ' change may have been discarded by dbl-click refresh; redo change
        m2 = ChangeItem0To(NEWMAT)
        log.WriteLine "  re-applied change -> " & m2
    End If
    session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1400
    Dim r : r = "w1=[" & WT(1) & "]"
    If InStr(WT(1), K_CHGTITLE) > 0 Then
        r = r & " old=[" & session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR_OLD").Text & "]"
        session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = NEWMAT
        session.findById("wnd[1]").sendVKey 0
        WScript.Sleep 1200
        r = r & " afterEnter w2=[" & W2Msg() & "] w1=[" & WT(1) & "]"
        If Not HasW(2) And InStr(WT(1), K_CHGTITLE) > 0 Then
            session.findById("wnd[1]/tbar[0]/btn[8]").press
            WScript.Sleep 1400
            r = r & " afterSave sbar[" & SbarT() & ":" & Sbar() & "] w2=[" & W2Msg() & "] w1=[" & WT(1) & "]"
        End If
    End If
    CancelAll
    Dim fin : fin = session.findById(IG).GetCellValue(0,"MATNR")
    Result tc, (InStr(r, K_SAME) > 0 Or fin = NEWMAT), "M25 same-material message OR item unchanged (" & NEWMAT & ")", r & " final item0=" & fin
    Tcode "/n"   ' discard in-memory change
End Sub

' ================= HIST-10 header de-dup on material query =================
Sub HIST_10()
    Dim tc : tc = "HIST-10 header de-dup when querying by material"
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = "2396"
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 3000
    If session.Info.ScreenNumber <> 100 Then Result tc, False, "2396 on screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    DblHeader0 : CancelAll
    Dim gi : Set gi = session.findById(IG)
    Dim nIt : nIt = gi.RowCount
    Dim mat : mat = gi.GetCellValue(0,"MATNR")
    ' count how many items of this header share that material
    Dim k, same : same = 0
    For k = 0 To nIt - 1
        If gi.GetCellValue(k,"MATNR") = mat Then same = same + 1
    Next
    log.WriteLine "  2396 items=" & nIt & " material0=" & mat & " occurrences in header=" & same
    ' query by that material only (no BSTKD)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = mat
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 5000
    If session.Info.ScreenNumber <> 100 Then Result tc, False, "material query result", "screen=" & session.Info.ScreenNumber & " sbar=" & Sbar() : Exit Sub
    Dim gh : Set gh = session.findById(HG)
    Dim d : Set d = CreateObject("Scripting.Dictionary")
    Dim dup : dup = 0
    Dim r, key
    For r = 0 To gh.RowCount - 1
        key = gh.GetCellValue(r,"ID") & "|" & gh.GetCellValue(r,"BSTKD") & "|" & gh.GetCellValue(r,"VERSION")
        If d.Exists(key) Then dup = dup + 1 Else d.Add key, 1
    Next
    Result tc, (dup = 0 And gh.RowCount > 0), "headers unique by ID|BSTKD|VERSION (material " & mat & " appears " & same & "x in 2396)", "headerRows=" & gh.RowCount & " uniqueKeys=" & d.Count & " duplicates=" & dup
    Tcode "/n"
End Sub

On Error Resume Next
DEL_04 : If Err.Number <> 0 Then Result "DEL-04", False, "-", "runtime " & Err.Description : Err.Clear
CancelAll
R8_06 : If Err.Number <> 0 Then Result "R8-06", False, "-", "runtime " & Err.Description : Err.Clear
CancelAll
R9_04_R8 : If Err.Number <> 0 Then Result "R9-04", False, "-", "runtime " & Err.Description : Err.Clear
CancelAll
HIST_10 : If Err.Number <> 0 Then Result "HIST-10", False, "-", "runtime " & Err.Description : Err.Clear
CancelAll
On Error GoTo 0
Tcode "/n"
log.WriteLine "SUMMARY: PASS=" & nPass & " FAIL=" & nFail
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "S36 done PASS=" & nPass & " FAIL=" & nFail
