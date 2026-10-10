' S33_r9_01_verify.vbs : R9-01 + UI-02 - change material and read item ALV in-place (NO re-query).
'   Material-change Save is in-memory only (no DB write until CREA_ORD), so /n discards it. Zero DB footprint.
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder("C:\sap-test\results\shots")
Set log = fso.CreateTextFile("C:\sap-test\logs\S33_r9_01_verify.txt", True, True)
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
log.WriteLine "System=" & session.Info.SystemName & " User=" & session.Info.User & " RunAt=" & Now
Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"
Dim K_CHGTITLE : K_CHGTITLE = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HCF54)&ChrW(&HB4DC)

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Function Sbar() : Sbar = session.findById("wnd[0]/sbar").Text : End Function
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

If Not OpenBSTKD("2190") Then log.WriteLine "abort: 2190 not on 100" : log.Close : WScript.Quit 10
Dim gh : Set gh = session.findById(HG)
Dim hr, useR : useR = -1
For hr = 0 To gh.RowCount - 1
    If gh.GetCellValue(hr,"VBELN")="" And gh.GetCellValue(hr,"STATU")<>"C" Then
        SelHeader hr
        If session.findById(IG).RowCount >= 1 Then
            If session.findById(IG).GetCellValue(0,"MATNR")="ML00007621" Then useR = hr : Exit For
        End If
    End If
Next
If useR < 0 Then log.WriteLine "abort: target not found" : log.Close : WScript.Quit 10
Dim beforeM, beforeMak, beforeU
beforeM = session.findById(IG).GetCellValue(0,"MATNR")
beforeMak = session.findById(IG).GetCellValue(0,"MAKTX")
beforeU = session.findById(IG).GetCellValue(0,"MEINS")
log.WriteLine "R9-01 target row=" & useR & " before: MATNR=" & beforeM & " MAKTX=[" & beforeMak & "] MEINS=" & beforeU

' open CHANGE_MAT, enter new material, Enter, Save
session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
session.findById(IG).pressToolbarButton "CHANGE_MAT"
WScript.Sleep 1500
If InStr(WT(1), K_CHGTITLE)=0 Then log.WriteLine "abort: dialog not open w1=" & WT(1) : SafeCancelAll : Tcode "/n" : log.Close : WScript.Quit 10
session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = "PNGB3500035L001"
session.findById("wnd[1]").sendVKey 0
WScript.Sleep 1400
Dim dlgMak : dlgMak = session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MAKTX").Text
log.WriteLine "dialog resolved MAKTX=[" & dlgMak & "] w2=[" & WT(2) & "]"
session.findById("wnd[1]/tbar[0]/btn[8]").press   ' Save (in-memory)
WScript.Sleep 1600
' answer Yes on any save-confirm (authorized; in-memory only)
Dim g2 : g2 = 0
Do While HasW(1) And g2 < 3
    Dim py, usr2, c2, comp2 : py = False
    On Error Resume Next
    Set usr2 = session.findById("wnd[1]/usr")
    If Err.Number = 0 Then
        For c2 = 0 To usr2.Children.Count-1
            Set comp2 = usr2.Children.ElementAt(c2)
            If comp2.Type = "GuiButton" Then
                If InStr(comp2.Text & comp2.Tooltip, ChrW(&HC608))>0 Then comp2.press : py=True : Exit For
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
    If Not py Then Exit Do
    WScript.Sleep 1400
    g2 = g2 + 1
Loop
log.WriteLine "after Save: screen=" & session.Info.ScreenNumber & " sbar=[" & Sbar() & "] w1=[" & WT(1) & "] w2=[" & WT(2) & "]"
SafeCancelAll

' READ ITEM ALV IN PLACE (no re-query)
Dim afterM, afterMak, afterU
On Error Resume Next
afterM = session.findById(IG).GetCellValue(0,"MATNR")
afterMak = session.findById(IG).GetCellValue(0,"MAKTX")
afterU = session.findById(IG).GetCellValue(0,"MEINS")
Err.Clear
On Error GoTo 0
log.WriteLine "R9-01 in-place after: MATNR=" & afterM & " MAKTX=[" & afterMak & "] MEINS=" & afterU
Dim ok : ok = (afterM = "PNGB3500035L001")
If ok Then log.WriteLine "[PASS] R9-01/UI-02 material change reflected in item ALV without re-query" _
    Else log.WriteLine "[CHECK] R9-01 item ALV MATNR did not become PNGB3500035L001"
' screenshot evidence
On Error Resume Next
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\R9_01_after_change.png", 2
Err.Clear
On Error GoTo 0

' discard in-memory change (do NOT press CREA_ORD): navigate away
Tcode "/n"
log.WriteLine "discarded via /n (no CREA_ORD, no DB write). done."
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "S33 done ok=" & ok
