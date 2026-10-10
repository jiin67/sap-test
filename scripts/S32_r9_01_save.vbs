' S32_r9_01_save.vbs : R9-01 designated positive - change ML00007621 -> PNGB3500035L001 on BSTKD 2190, SAVE.
'   Captures success msg (M29), re-query result, and Z33SDT644 / Z33SDT641 rows for the target ID.
'   NO cleanup here (delete handled separately once exact key is known). Authorized: presses Yes on save-confirm.
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S32_r9_01_save.txt", True, True)
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
Dim K_CHGTITLE, K_YES, K_UNIQ
K_CHGTITLE = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HCF54)&ChrW(&HB4DC)   ' jaejae-code
K_YES      = ChrW(&HC608)                                          ' Yes
K_UNIQ     = "Uniq"

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
Sub SafeCancelAll()
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
                            If InStr(comp.Text & comp.Tooltip, ChrW(&HC544)&ChrW(&HB2C8)&ChrW(&HC624))>0 Or InStr(comp.Text & comp.Tooltip, ChrW(&HCDE8)&ChrW(&HC18C))>0 Then comp.press : did=True : Exit For
                        End If
                    Next
                End If
                Err.Clear
                If Not did Then session.findById("wnd[" & n & "]/tbar[0]/btn[12]").press : Err.Clear
                If HasW(n) Then session.findById("wnd[" & n & "]").close : Err.Clear
                WScript.Sleep 500
            End If
        Next
    Next
    On Error GoTo 0
End Sub
Function OpenBSTKD(b)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = b
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 3500
    OpenBSTKD = (session.Info.ScreenNumber = 100)
End Function
' select header row r, populate items
Sub SelHeader(r)
    Dim gh : Set gh = session.findById(HG)
    gh.setCurrentCell r,"BSTKD" : gh.selectedRows = CStr(r)
    On Error Resume Next
    gh.doubleClickCurrentCell
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 1500
    SafeCancelAll
End Sub
' SE16N read: return concatenated rows of `table` where ID column = idv (filter in VBS)
Sub Se16nDump(table, idv, label)
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    ' ensure edit off
    If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "9999"
    ' set ID condition: find condition row whose desc contains 'Uniq'
    Dim tc, rr, desc, setId
    setId = False
    Set tc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC")
    For rr = 0 To 20
        desc = ""
        desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, K_UNIQ) > 0 Then
            session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = idv
            setId = True
            Exit For
        End If
    Next
    log.WriteLine "  [SE16N " & label & "] table=" & table & " idCondSet=" & setId
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2500
    ' read result ALV
    Dim g, cols, r2, ln, cc
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then
        log.WriteLine "  (no result ALV) screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " sbar=" & Sbar()
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    log.WriteLine "  rows=" & g.RowCount
    Set cols = g.ColumnOrder
    For r2 = 0 To g.RowCount - 1
        If r2 > 20 Then Exit For
        ln = "   #" & r2 & ":"
        For cc = 0 To cols.Count - 1
            Dim cid : cid = cols(cc)
            If cid="ID" Or cid="BSTKD" Or cid="VERSION" Or cid="SEQNO" Or cid="MATNR" Or cid="MATNR_OLD" Or cid="MAKTX" Or cid="MEINS" Or cid="VBELN" Or cid="ERNAM" Or cid="ERDAT" Or cid="ERZET" Or cid="LGORT" Then
                ln = ln & " " & cid & "=[" & g.GetCellValue(r2, cid) & "]"
            End If
        Next
        log.WriteLine ln
    Next
    On Error GoTo 0
End Sub

' ===== R9-01 =====
Dim targetID, targetVER, targetSEQ, beforeMat
targetID = "" : beforeMat = ""
If Not OpenBSTKD("2190") Then
    log.WriteLine "FATAL: 2190 not on screen 100 (screen=" & session.Info.ScreenNumber & ")"
    log.Close : WScript.Echo "S32 abort" : WScript.Quit 10
End If
Dim gh : Set gh = session.findById(HG)
Dim hr, useR
useR = -1
For hr = 0 To gh.RowCount - 1
    If gh.GetCellValue(hr,"VBELN") = "" And gh.GetCellValue(hr,"STATU") <> "C" Then
        SelHeader hr
        If session.findById(IG).RowCount >= 1 Then
            If session.findById(IG).GetCellValue(0,"MATNR") = "ML00007621" Then
                useR = hr
                targetID = gh.GetCellValue(hr,"ID")
                targetVER = gh.GetCellValue(hr,"VERSION")
                targetSEQ = session.findById(IG).GetCellValue(0,"SEQNO")
                beforeMat = session.findById(IG).GetCellValue(0,"MATNR")
                Exit For
            End If
        End If
    End If
Next
log.WriteLine "TARGET header row=" & useR & " ID=" & targetID & " VERSION=" & targetVER & " SEQNO=" & targetSEQ & " MATNR(before)=" & beforeMat
If useR < 0 Then log.WriteLine "FATAL: target not found" : log.Close : WScript.Echo "S32 abort" : WScript.Quit 10

' baseline Z33SDT644 for this ID
log.WriteLine "=== BASELINE Z33SDT644 (ID=" & targetID & ") ==="
Se16nDump "Z33SDT644", targetID, "baseline644"

' perform change + save
If Not OpenBSTKD("2190") Then log.WriteLine "re-open fail" : log.Close : WScript.Quit 10
SelHeader useR
session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
session.findById(IG).pressToolbarButton "CHANGE_MAT"
WScript.Sleep 1500
If InStr(WT(1), K_CHGTITLE) = 0 Then
    log.WriteLine "FATAL: change dialog did not open. w1=" & WT(1) & " sbar=" & Sbar()
    SafeCancelAll : Tcode "/n" : log.Close : WScript.Echo "S32 abort" : WScript.Quit 10
End If
session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = "PNGB3500035L001"
session.findById("wnd[1]").sendVKey 0
WScript.Sleep 1500
log.WriteLine "after Enter: newMAKTX=[" & session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MAKTX").Text & "] w2=[" & WT(2) & "]"
' Save
session.findById("wnd[1]/tbar[0]/btn[8]").press
WScript.Sleep 1800
log.WriteLine "after Save: sbar[" & SbarT() & ":" & Sbar() & "] w1=[" & WT(1) & "] w2=[" & WT(2) & "]"
' handle save-confirm Yes/No -> Yes (authorized)
Dim guard2 : guard2 = 0
Do While HasW(1) And guard2 < 4
    Dim pressedYes, usr2, c2, comp2
    pressedYes = False
    On Error Resume Next
    Set usr2 = session.findById("wnd[1]/usr")
    If Err.Number = 0 Then
        For c2 = 0 To usr2.Children.Count - 1
            Set comp2 = usr2.Children.ElementAt(c2)
            If comp2.Type = "GuiButton" Then
                If InStr(comp2.Text & comp2.Tooltip, K_YES) > 0 Then comp2.press : pressedYes = True : Exit For
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
    If pressedYes Then
        WScript.Sleep 1800
        log.WriteLine "  pressed YES on confirm -> sbar[" & SbarT() & ":" & Sbar() & "] w1=[" & WT(1) & "] w2=[" & WT(2) & "]"
    Else
        Exit Do
    End If
    guard2 = guard2 + 1
Loop
' if an error popup (오류) remains, log and close
If HasW(2) Then
    On Error Resume Next
    log.WriteLine "  wnd2 msg=[" & session.findById("wnd[2]/usr/txtMESSTXT1").Text & "]"
    Err.Clear
    On Error GoTo 0
End If
SafeCancelAll
log.WriteLine "post-save screen=" & session.Info.ScreenNumber & " sbar=[" & Sbar() & "]"

' re-query ZSD600 to see applied material
Dim afterMat, afterMak
afterMat = "?" : afterMak = "?"
If OpenBSTKD("2190") Then
    Dim gh3 : Set gh3 = session.findById(HG)
    Dim hh
    For hh = 0 To gh3.RowCount - 1
        If gh3.GetCellValue(hh,"ID") = targetID Then
            SelHeader hh
            If session.findById(IG).RowCount >= 1 Then
                afterMat = session.findById(IG).GetCellValue(0,"MATNR")
                afterMak = session.findById(IG).GetCellValue(0,"MAKTX")
            End If
            Exit For
        End If
    Next
End If
log.WriteLine "REQUERY item MATNR(after)=[" & afterMat & "] MAKTX=[" & afterMak & "]"

' Z33SDT644 after
log.WriteLine "=== AFTER Z33SDT644 (ID=" & targetID & ") ==="
Se16nDump "Z33SDT644", targetID, "after644"
' Z33SDT641 after (did item table change?)
log.WriteLine "=== AFTER Z33SDT641 (ID=" & targetID & ") ==="
Se16nDump "Z33SDT641", targetID, "after641"

Tcode "/n"
log.WriteLine "done. TARGET ID=" & targetID & " VERSION=" & targetVER & " SEQNO=" & targetSEQ
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "S32 done"
