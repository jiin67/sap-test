' probe_buttons.vbs : find out what the result-screen buttons btn[26] / btn[27] do.
'   READ ONLY : presses the buttons, records the popup, closes it.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

Set connection = Nothing
For i = 0 To application.Children.Count - 1
    Set base = application.Children.ElementAt(i).Children.ElementAt(0)
    If base.Info.SystemName = "KDQ" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then
    WScript.Echo "GUARD: no KDQ connection"
    WScript.Quit 99
End If
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then
    WScript.Echo "GUARD: " & base.Info.SystemName
    WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\probe_buttons.txt", True, True)

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
log.WriteLine "work session=" & session.Id

Function SetField(nm, val)
    Dim o
    On Error Resume Next
    Set o = session.findById("wnd[0]/usr/ctxt" & nm)
    If Err.Number <> 0 Then
        Err.Clear
        Set o = session.findById("wnd[0]/usr/txt" & nm)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        SetField = False
        Exit Function
    End If
    o.Text = val
    SetField = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Sub DumpWnd(w, tag, depth)
    Dim c, k, comp, t, sb
    On Error Resume Next
    k = w.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For c = 0 To k - 1
        Set comp = w.Children.ElementAt(c)
        t = comp.Type
        If t = "GuiShell" Then
            sb = comp.SubType
            log.WriteLine tag & " SHELL " & Mid(comp.Id, InStr(comp.Id, "/wnd[")) & " sub=" & sb
            If sb = "GridView" Then
                Dim ord, r, cc, line
                Set ord = comp.ColumnOrder
                line = ""
                For cc = 0 To ord.Count - 1
                    If cc > 0 Then line = line & " | "
                    line = line & ord(cc)
                Next
                log.WriteLine tag & " COLS " & line
                For r = 0 To comp.RowCount - 1
                    line = ""
                    For cc = 0 To ord.Count - 1
                        If cc > 0 Then line = line & " | "
                        line = line & comp.GetCellValue(r, ord(cc))
                    Next
                    log.WriteLine tag & " ROW" & r & " " & line
                    Err.Clear
                    If r > 25 Then Exit For
                Next
            End If
            Err.Clear
        ElseIf t = "GuiLabel" Or t = "GuiTextField" Or t = "GuiCTextField" Or t = "GuiButton" Then
            If Trim(comp.Text) <> "" Then
                log.WriteLine tag & " " & t & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[")) & " [" & comp.Text & "]"
            End If
            Err.Clear
        End If
        DumpWnd comp, tag, depth + 1
    Next
    Err.Clear
End Sub

Sub ClosePopup()
    Dim w, g
    g = 0
    Do
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        On Error GoTo 0
        w.sendVKey 12
        WScript.Sleep 900
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number = 0 Then
            w.close
            WScript.Sleep 700
        End If
        Err.Clear
        On Error GoTo 0
        g = g + 1
    Loop While g < 3
End Sub

' ---- open the result screen with a small result set ----
Dim ok
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2200
session.findById("wnd[0]/usr/radPA_R01").Select
session.findById("wnd[0]/usr/radPA_CHX").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
Err.Clear
On Error GoTo 0
ok = SetField("SO_RCDT-LOW", "2026.07.01")
ok = SetField("SO_RCDT-HIGH", "2026.07.01")
ok = SetField("SO_CSLTN-LOW", "2026051400858")
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2500
Err.Clear
log.WriteLine "result screen=" & session.Info.ScreenNumber & " title=[" & session.findById("wnd[0]").Text & "]"
Err.Clear
On Error GoTo 0

Dim btns, bi, bn
btns = Array(26, 27)
For bi = 0 To UBound(btns)
    bn = btns(bi)
    log.WriteLine "================ tbar[1]/btn[" & bn & "] ================"
    On Error Resume Next
    Dim b
    Set b = session.findById("wnd[0]/tbar[1]/btn[" & bn & "]")
    If Err.Number <> 0 Then
        log.WriteLine "  button not found"
        Err.Clear
    Else
        log.WriteLine "  text=[" & b.Text & "] tooltip=[" & b.Tooltip & "]"
        b.press
        WScript.Sleep 2200
        Err.Clear
        Dim w
        Set w = session.findById("wnd[1]")
        If Err.Number = 0 Then
            log.WriteLine "  POPUP title=[" & w.Text & "]"
            w.HardCopy "C:\sap-test\results\shots\MANUAL_btn" & bn & ".png", 2
            Err.Clear
            DumpWnd w, "  btn" & bn, 1
            ClosePopup
        Else
            Err.Clear
            log.WriteLine "  no popup. screen=" & session.Info.ScreenNumber & _
                          " title=[" & session.findById("wnd[0]").Text & "]" & _
                          " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
            session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\MANUAL_btn" & bn & ".png", 2
            Err.Clear
            ' if it navigated away, go back
            If session.Info.ScreenNumber <> 100 Then
                session.findById("wnd[0]").sendVKey 3
                WScript.Sleep 1800
                Err.Clear
            End If
        End If
    End If
    On Error GoTo 0
Next

' capture the selection screen and the result screen for the manual
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
ClosePopup
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2200
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\MANUAL_sel_center.png", 2
session.findById("wnd[0]/usr/radPA_R02").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1300
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\MANUAL_sel_dealer.png", 2
session.findById("wnd[0]/usr/radPA_R01").Select
session.findById("wnd[0]/usr/radPA_NCH").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1300
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\MANUAL_sel_notrecv.png", 2
Err.Clear
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
ClosePopup
Err.Clear
On Error GoTo 0

log.WriteLine "done"
log.Close
On Error Resume Next
session.findById("wnd[0]").close
Err.Clear
On Error GoTo 0
WScript.Echo "probe_buttons done"
