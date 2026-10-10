' probe_dropdown.vbs : work out the accepted input format for the 부적합사유 cell,
'   and learn how the program reacts when leaving with unsaved changes.
'   Enters values into the grid but NEVER saves - on exit it explicitly presses
'   the "아니오 / No" button of any save-confirmation popup.
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
Set log = fso.CreateTextFile("C:\sap-test\logs\probe_dropdown.txt", True, True)

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

' dump a popup and press the button whose text means "No"
Sub HandlePopupNo(tag)
    Dim w, j, b, txt, pressed
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        log.WriteLine tag & " : no popup"
        Exit Sub
    End If
    log.WriteLine tag & " : POPUP [" & w.Text & "]"
    ' log every label and button we can see
    DumpPopup w, 1
    pressed = False
    For j = 0 To 8
        Set b = session.findById("wnd[1]/usr/btnSPOP-OPTION" & j)
        If Err.Number = 0 Then
            txt = b.Text
            log.WriteLine "   btnSPOP-OPTION" & j & " = [" & txt & "]"
            If InStr(txt, "아니") > 0 Or InStr(UCase(txt), "NO") > 0 Then
                b.press
                WScript.Sleep 1200
                log.WriteLine "   pressed NO"
                pressed = True
                Exit For
            End If
        End If
        Err.Clear
    Next
    If Not pressed Then
        ' fall back to the standard cancel key, never Enter
        Set w = session.findById("wnd[1]")
        If Err.Number = 0 Then
            w.sendVKey 12
            WScript.Sleep 1200
            log.WriteLine "   sent F12 (cancel)"
        End If
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub DumpPopup(root, depth)
    Dim c, k, comp, t
    On Error Resume Next
    k = root.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For c = 0 To k - 1
        Set comp = root.Children.ElementAt(c)
        t = comp.Type
        If t = "GuiLabel" Or t = "GuiButton" Or t = "GuiTextField" Then
            log.WriteLine String(depth * 2, " ") & t & " " & _
                          Mid(comp.Id, InStr(comp.Id, "/wnd[1]")) & " [" & comp.Text & "]"
        End If
        Err.Clear
        DumpPopup comp, depth + 1
    Next
    Err.Clear
End Sub

' ---------- open 미접수 restricted to the one 상담번호 ----------
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
log.WriteLine "rows=" & g.RowCount
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
log.WriteLine "target row=" & target

' ---------- try candidate values for the 부적합사유 cell ----------
Dim cands, ci, v
cands = Array("Z02 회수기간초과", "회수기간초과", "Z02")
For ci = 0 To UBound(cands)
    On Error Resume Next
    g.setCurrentCell target, "NCF_RSN_VT"
    g.modifyCell target, "NCF_RSN_VT", cands(ci)
    If Err.Number <> 0 Then
        log.WriteLine "candidate [" & cands(ci) & "] modifyCell ERR " & Err.Description
        Err.Clear
    Else
        v = g.GetCellValue(target, "NCF_RSN_VT")
        log.WriteLine "candidate [" & cands(ci) & "] -> readback [" & v & "]"
        Err.Clear
    End If
    On Error GoTo 0
Next

' also record what an F4 on that cell offers
On Error Resume Next
g.setCurrentCell target, "NCF_RSN_VT"
session.findById("wnd[0]").sendVKey 4
WScript.Sleep 1500
Err.Clear
On Error GoTo 0
Dim w2
On Error Resume Next
Set w2 = session.findById("wnd[1]")
If Err.Number = 0 Then
    log.WriteLine "F4 POPUP [" & w2.Text & "]"
    DumpPopup w2, 1
    w2.sendVKey 12
    WScript.Sleep 900
End If
Err.Clear
On Error GoTo 0

' ---------- also check the 책임사업장 F4 (TC-10) ----------
On Error Resume Next
g.setCurrentCell target, "RSPC"
session.findById("wnd[0]").sendVKey 4
WScript.Sleep 1500
Err.Clear
Set w2 = session.findById("wnd[1]")
If Err.Number = 0 Then
    log.WriteLine "RSPC F4 POPUP [" & w2.Text & "]"
    DumpPopup w2, 1
    w2.sendVKey 12
    WScript.Sleep 900
End If
Err.Clear
On Error GoTo 0

' ---------- leave WITHOUT saving ----------
log.WriteLine "leaving without saving (F3)"
On Error Resume Next
session.findById("wnd[0]").sendVKey 3
WScript.Sleep 2000
Err.Clear
On Error GoTo 0
HandlePopupNo "after F3"

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
Err.Clear
On Error GoTo 0
HandlePopupNo "after /n"

log.WriteLine "final prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber
log.WriteLine "done"
log.Close
On Error Resume Next
session.findById("wnd[0]").close
Err.Clear
On Error GoTo 0
WScript.Echo "probe_dropdown done"
