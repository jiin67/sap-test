' S31_chgmat_probe.vbs : probe material-change dialog behaviour on Enter (READ ONLY - always Cancel, never Save)
'   - CHANGE_MAT with 0 selected / 2 selected  (capture guard message: sbar + wnd2 popup)
'   - dialog: enter PNGB3500035L001 + Enter -> does MAKTX resolve? any error/ATP popup?
'   - dialog: enter same material + Enter -> message?
'   Also probe DEL_DELIVERY confirm dialog buttons.
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S31_chgmat_probe.txt", True, True)
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
Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Function Sbar() : Sbar = session.findById("wnd[0]/sbar").Text : End Function
Function HasW(n)
    On Error Resume Next
    HasW = False
    Dim w : Set w = session.findById("wnd[" & n & "]")
    If Err.Number = 0 Then HasW = True
    Err.Clear
    On Error GoTo 0
End Function
Function WT(n)
    On Error Resume Next
    WT = "" : WT = session.findById("wnd[" & n & "]").Text
    Err.Clear
    On Error GoTo 0
End Function
Sub Dump(comp, depth)
    Dim c, k, line
    If comp.Type = "GuiMenubar" Then Exit Sub
    line = String(depth*2," ") & comp.Type
    On Error Resume Next
    line = line & "/" & comp.SubType
    On Error GoTo 0
    line = line & " | " & comp.Id
    On Error Resume Next
    line = line & " | [" & Left(comp.Text,80) & "]"
    On Error GoTo 0
    log.WriteLine line
    On Error Resume Next
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k-1
        Dump comp.Children.ElementAt(c), depth+1
    Next
End Sub
Sub SafeCancelAll()
    Dim g, n, usr, c, comp, closed
    On Error Resume Next
    For g = 1 To 10
        closed = True
        For n = 3 To 1 Step -1
            If HasW(n) Then
                closed = False
                ' prefer No/Cancel button
                Dim did : did = False
                Set usr = session.findById("wnd[" & n & "]/usr")
                If Err.Number = 0 Then
                    For c = 0 To usr.Children.Count-1
                        Set comp = usr.Children.ElementAt(c)
                        If comp.Type = "GuiButton" Then
                            If InStr(comp.Text & comp.Tooltip, ChrW(&HC544)&ChrW(&HB2C8)&ChrW(&HC624))>0 Or InStr(comp.Text & comp.Tooltip, ChrW(&HCDE8)&ChrW(&HC18C))>0 Then comp.press : did = True : Exit For
                        End If
                    Next
                End If
                Err.Clear
                If Not did Then session.findById("wnd[" & n & "]/tbar[0]/btn[12]").press : Err.Clear
                If HasW(n) Then session.findById("wnd[" & n & "]").close : Err.Clear
                WScript.Sleep 600
            End If
        Next
        If closed Then Exit For
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
Sub SelHeaderWithItem(matnr)
    Dim gh, hr
    Set gh = session.findById(HG)
    For hr = 0 To gh.RowCount-1
        gh.setCurrentCell hr,"BSTKD" : gh.selectedRows = CStr(hr)
        On Error Resume Next
        gh.doubleClickCurrentCell
        Err.Clear
        On Error GoTo 0
        WScript.Sleep 1300
        SafeCancelAll
        If session.findById(IG).RowCount >= 1 Then
            If matnr = "" Or session.findById(IG).GetCellValue(0,"MATNR") = matnr Then Exit For
        End If
    Next
End Sub

' ---- Probe 1: CHANGE_MAT with 0 selected on 2190 ----
log.WriteLine "=== P1: CHANGE_MAT 0 selected (2190) ==="
If OpenBSTKD("2190") Then
    SelHeaderWithItem "ML00007621"
    session.findById(IG).selectedRows = ""
    session.findById(IG).currentCellRow = -1
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 800
    log.WriteLine "  immediate sbar=[" & Sbar() & "] w1=[" & WT(1) & "] w2=[" & WT(2) & "]"
    If HasW(2) Then log.WriteLine "  w2 dump:" : Dump session.findById("wnd[2]"), 1
    SafeCancelAll
End If

' ---- Probe 2: CHANGE_MAT with 2 selected on 2396 ----
log.WriteLine "=== P2: CHANGE_MAT 2 selected (2396) ==="
If OpenBSTKD("2396") Then
    SelHeaderWithItem ""
    Dim ic : ic = session.findById(IG).RowCount
    log.WriteLine "  2396 item rows=" & ic
    If ic >= 2 Then
        session.findById(IG).selectedRows = "0,1"
        session.findById(IG).pressToolbarButton "CHANGE_MAT"
        WScript.Sleep 800
        log.WriteLine "  immediate sbar=[" & Sbar() & "] w1=[" & WT(1) & "] w2=[" & WT(2) & "]"
        If HasW(2) Then log.WriteLine "  w2 dump:" : Dump session.findById("wnd[2]"), 1
        SafeCancelAll
    End If
End If

' ---- Probe 3: dialog enter PNGB3500035L001 + Enter (2190) ----
log.WriteLine "=== P3: enter new material PNGB3500035L001 + Enter (no save) ==="
If OpenBSTKD("2190") Then
    SelHeaderWithItem "ML00007621"
    session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1300
    log.WriteLine "  dialog=[" & WT(1) & "]"
    If InStr(WT(1), ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HCF54)&ChrW(&HB4DC)) > 0 Then
        log.WriteLine "  --- dialog structure before entry ---"
        Dump session.findById("wnd[1]/usr"), 1
        session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = "PNGB3500035L001"
        session.findById("wnd[1]").sendVKey 0
        WScript.Sleep 1500
        log.WriteLine "  after Enter: w1=[" & WT(1) & "] w2=[" & WT(2) & "] mainSbar=[" & Sbar() & "]"
        On Error Resume Next
        log.WriteLine "  new MATNR field=[" & session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text & "] MAKTX=[" & session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MAKTX").Text & "]"
        Err.Clear
        On Error GoTo 0
        If HasW(2) Then log.WriteLine "  w2 dump:" : Dump session.findById("wnd[2]"), 1
        log.WriteLine "  dialog tbar buttons:"
        Dim tb, c
        Set tb = session.findById("wnd[1]/tbar[0]")
        For c = 0 To tb.Children.Count-1
            log.WriteLine "    " & tb.Children.ElementAt(c).Id & " [" & tb.Children.ElementAt(c).Text & "] tip=[" & tb.Children.ElementAt(c).Tooltip & "]"
        Next
    End If
    SafeCancelAll
End If

' ---- Probe 4: dialog enter SAME material + Enter (2190) ----
log.WriteLine "=== P4: enter same material ML00007621 + Enter (no save) ==="
If OpenBSTKD("2190") Then
    SelHeaderWithItem "ML00007621"
    session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1300
    If InStr(WT(1), ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HCF54)&ChrW(&HB4DC)) > 0 Then
        session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = "ML00007621"
        session.findById("wnd[1]").sendVKey 0
        WScript.Sleep 1500
        log.WriteLine "  after Enter same: w1=[" & WT(1) & "] w2=[" & WT(2) & "] mainSbar=[" & Sbar() & "]"
        If HasW(2) Then log.WriteLine "  w2 dump:" : Dump session.findById("wnd[2]"), 1
    End If
    SafeCancelAll
End If

' ---- Probe 5: DEL_DELIVERY confirm dialog on 2190 (no delivery) ----
log.WriteLine "=== P5: DEL_DELIVERY confirm dialog buttons (2190, no delivery) ==="
If OpenBSTKD("2190") Then
    Dim gh2 : Set gh2 = session.findById(HG)
    gh2.setCurrentCell 0,"BSTKD" : gh2.selectedRows = "0"
    gh2.pressToolbarButton "DEL_DELIVERY"
    WScript.Sleep 2000
    log.WriteLine "  w1=[" & WT(1) & "]"
    If HasW(1) Then Dump session.findById("wnd[1]/usr"), 1
    SafeCancelAll
End If

Tcode "/n"
log.WriteLine "done"
log.Close
Dim leftover : leftover = HasW(1)
If Not leftover Then session.findById("wnd[0]").close
WScript.Echo "S31 probe done leftoverPopup=" & leftover
