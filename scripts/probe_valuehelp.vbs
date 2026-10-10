' probe_valuehelp.vbs : the result-screen buttons 부적합 사유 (Ctrl+F2) and
'   책임사업장 (Ctrl+F3) act on the currently selected ALV cell. Select an
'   editable row first, then press them and record the value list.
'   READ ONLY : never saves. Leaves the screen answering "No" to the save prompt.
'   args: <CSLTNO> <MATNR_PRT> <TAGNO1>
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

Dim CS, MT, TG, ARGS
Set ARGS = WScript.Arguments
CS = ARGS(0)
MT = ARGS(1)
TG = ARGS(2)

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\probe_valuehelp.txt", True, True)

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

Sub DumpWnd(w, tag)
    Dim c, k, comp, t, sb, ord, r, cc, line
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
                    If r > 30 Then Exit For
                Next
            End If
            Err.Clear
        ElseIf t = "GuiLabel" Or t = "GuiTextField" Or t = "GuiCTextField" Then
            If Trim(comp.Text) <> "" Then
                log.WriteLine tag & " " & t & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[")) & " [" & comp.Text & "]"
            End If
            Err.Clear
        End If
        DumpWnd comp, tag
    Next
    Err.Clear
End Sub

Sub ClosePopup(tag)
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

' ---- open 미접수 restricted to one consultation number ----
Dim ok
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2200
session.findById("wnd[0]/usr/radPA_R01").Select
session.findById("wnd[0]/usr/radPA_NCH").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
Err.Clear
On Error GoTo 0
ok = SetField("SO_RCDT-LOW", "")
ok = SetField("SO_RCDT-HIGH", "")
ok = SetField("SO_PRDRE-LOW", "2026.07.01")
ok = SetField("SO_PRDRE-HIGH", "2026.07.15")
ok = SetField("SO_CSLTN-LOW", CS)

On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2500
Err.Clear
On Error GoTo 0

Dim g, r, target
On Error Resume Next
Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
If Err.Number <> 0 Then
    Err.Clear
    log.WriteLine "no grid"
    log.Close
    WScript.Quit 10
End If
target = -1
For r = 0 To g.RowCount - 1
    If g.GetCellValue(r, "MATNR_PRT") = MT And g.GetCellValue(r, "TAGNO1") = TG Then
        target = r
        Exit For
    End If
    Err.Clear
Next
Err.Clear
On Error GoTo 0
If target < 0 Then
    log.WriteLine "target not found"
    log.Close
    WScript.Quit 11
End If
log.WriteLine "rows=" & g.RowCount & " target row=" & target

' ---- 부적합 사유 : select the NCF_RSN_VT cell, press Ctrl+F2 (btn[26]) ----
log.WriteLine "=========== 부적합 사유 (btn[26], Ctrl+F2) ==========="
On Error Resume Next
g.setCurrentCell target, "NCF_RSN_VT"
g.selectedRows = CStr(target)
session.findById("wnd[0]/tbar[1]/btn[26]").press
WScript.Sleep 2200
Err.Clear
Dim w
Set w = session.findById("wnd[1]")
If Err.Number = 0 Then
    log.WriteLine "  POPUP [" & w.Text & "]"
    w.HardCopy "C:\sap-test\results\shots\MANUAL_ncfrsn_list.png", 2
    Err.Clear
    DumpWnd w, "  NCFRSN"
    ClosePopup "ncfrsn"
Else
    Err.Clear
    log.WriteLine "  no popup. sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    Err.Clear
End If
On Error GoTo 0

' ---- 책임사업장 : select the RSPC cell, press Ctrl+F3 (btn[27]) ----
log.WriteLine "=========== 책임사업장 (btn[27], Ctrl+F3) ==========="
On Error Resume Next
Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
g.setCurrentCell target, "RSPC"
g.selectedRows = CStr(target)
session.findById("wnd[0]/tbar[1]/btn[27]").press
WScript.Sleep 2200
Err.Clear
Set w = session.findById("wnd[1]")
If Err.Number = 0 Then
    log.WriteLine "  POPUP [" & w.Text & "]"
    w.HardCopy "C:\sap-test\results\shots\MANUAL_rspc_list.png", 2
    Err.Clear
    DumpWnd w, "  RSPC"
    ClosePopup "rspc"
Else
    Err.Clear
    log.WriteLine "  no popup. sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    Err.Clear
End If
On Error GoTo 0

' ---- leave discarding any change ----
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1800
Err.Clear
Dim b
Set b = session.findById("wnd[1]/usr/btnBUTTON_2")
If Err.Number = 0 Then
    log.WriteLine "exit popup -> pressing [" & b.Text & "]"
    b.press
    WScript.Sleep 1200
End If
Err.Clear
ClosePopup "exit"
On Error GoTo 0

log.WriteLine "final screen=" & session.Info.ScreenNumber
log.WriteLine "done"
log.Close
On Error Resume Next
session.findById("wnd[0]").close
Err.Clear
On Error GoTo 0
WScript.Echo "probe_valuehelp done"
