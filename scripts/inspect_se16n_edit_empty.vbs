' inspect_se16n_edit_empty.vbs : probe SE16N edit mode on ZCRM_CHRG_CNFRM (empty). Changes nothing, never saves.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, g, c
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_se16n_edit_empty.txt", True, True)
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" Then WScript.Quit 99

Function Exists(id)
    Dim o
    On Error Resume Next
    Set o = session.findById(id)
    Exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
Sub Status(t)
    On Error Resume Next
    log.WriteLine t & " prog=" & session.Info.Program & " screen=" & session.Info.ScreenNumber & " title=[" & session.findById("wnd[0]").Text & "] sbar=[" & session.findById("wnd[0]/sbar").MessageType & "] " & session.findById("wnd[0]/sbar").Text
    Err.Clear
    On Error GoTo 0
End Sub
Sub DumpTree(comp, depth, lvl)
    Dim k, t, s, j
    On Error Resume Next
    t = comp.Type
    s = String(lvl * 2, " ") & t & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[") + 1) & " text=[" & comp.Text & "]"
    If t = "GuiShell" Then s = s & " sub=" & comp.SubType
    log.WriteLine s
    Err.Clear
    If lvl >= depth Then Exit Sub
    k = comp.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    For j = 0 To k - 1
        DumpTree comp.Children.ElementAt(j), depth, lvl + 1
    Next
End Sub
Sub PopupCancel(where)
    If Exists("wnd[1]") Then
        log.WriteLine "POPUP (" & where & ") [" & session.findById("wnd[1]").Text & "]"
        DumpTree session.findById("wnd[1]"), 4, 0
        session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\SE16N_edit_empty_popup_" & where & ".png", 2
        session.findById("wnd[1]").sendVKey 12
        WScript.Sleep 1000
    End If
End Sub

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "ZCRM_CHRG_CNFRM"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "50"
session.findById("wnd[0]/usr/chkGD-EDIT").Selected = True
Err.Clear
On Error GoTo 0
Status "before F8:"
log.WriteLine "---- tbar[1] on selection screen"
For i = 0 To 50
    If Exists("wnd[0]/tbar[1]/btn[" & i & "]") Then log.WriteLine "  btn[" & i & "] tip=[" & session.findById("wnd[0]/tbar[1]/btn[" & i & "]").Tooltip & "]"
Next
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
Status "after F8:"
PopupCancel "f8"
Status "after popup:"
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\SE16N_edit_empty.png", 2
log.WriteLine "---- tbar[1]"
For i = 0 To 50
    If Exists("wnd[0]/tbar[1]/btn[" & i & "]") Then log.WriteLine "  btn[" & i & "] tip=[" & session.findById("wnd[0]/tbar[1]/btn[" & i & "]").Tooltip & "]"
Next
If Exists("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    On Error Resume Next
    log.WriteLine "GRID rows=" & g.RowCount & " cols=" & g.ColumnCount
    For c = 0 To g.ToolbarButtonCount - 1
        log.WriteLine "  tb[" & c & "] id=" & g.GetToolbarButtonId(c) & " type=" & g.GetToolbarButtonType(c) & " enabled=" & g.GetToolbarButtonEnabled(c) & " tip=[" & g.GetToolbarButtonTooltip(c) & "]"
        Err.Clear
    Next
    Dim ord
    Set ord = g.ColumnOrder
    For c = 0 To ord.Count - 1
        log.WriteLine "  col " & ord(c)
    Next
    Err.Clear
    On Error GoTo 0
Else
    log.WriteLine "no RESULT_LIST grid - usr tree:"
    DumpTree session.findById("wnd[0]/usr"), 3, 0
End If
' leave without saving
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
Err.Clear
On Error GoTo 0
PopupCancel "leave"
On Error Resume Next
session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
log.WriteLine "edit checkbox final=" & session.findById("wnd[0]/usr/chkGD-EDIT").Selected
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
Err.Clear
On Error GoTo 0
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Quit 0