' inspect_se37_table.vbs : READ-ONLY. SE37 single test of the RFC -> open the IT_DATA table editor and dump it.
'   Never presses F8 on the test screen (the function module is NOT executed).
'   log: logs\inspect_se37_table.txt (UTF-16), shot: results\shots\SE37_it_data_editor.png
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, try

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_se37_table.txt", True, True)

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" And session.Info.SystemName <> "KDD" Then WScript.Quit 99

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
    log.WriteLine t & " prog=" & session.Info.Program & " screen=" & session.Info.ScreenNumber & _
        " title=[" & session.findById("wnd[0]").Text & "] sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    Err.Clear
    On Error GoTo 0
End Sub

Sub DumpTree(comp, depth, lvl)
    Dim c, k, t, s
    On Error Resume Next
    t = comp.Type
    s = String(lvl * 2, " ") & t & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[") + 1) & " text=[" & comp.Text & "]"
    If t = "GuiShell" Then s = s & " sub=" & comp.SubType
    If t = "GuiTableControl" Then s = s & " rows=" & comp.RowCount & " visible=" & comp.VisibleRowCount
    s = s & " tip=[" & comp.Tooltip & "] changeable=" & comp.Changeable
    log.WriteLine s
    Err.Clear
    If lvl >= depth Then Exit Sub
    k = comp.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    For c = 0 To k - 1
        DumpTree comp.Children.ElementAt(c), depth, lvl + 1
    Next
End Sub

Sub Tbar1()
    For i = 0 To 50
        If Exists("wnd[0]/tbar[1]/btn[" & i & "]") Then
            On Error Resume Next
            log.WriteLine "  btn[" & i & "] tip=[" & session.findById("wnd[0]/tbar[1]/btn[" & i & "]").Tooltip & "]"
            Err.Clear
            On Error GoTo 0
        End If
    Next
End Sub

' ---------- test screen ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE37"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtRS38L-NAME").Text = "ZCRMF03_RFC_ZCRM_CHRG_CNFRM"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
Status "test screen:"
If session.Info.Program <> "SAPMSSY0" Then
    log.WriteLine "not on test screen - abort"
    log.Close
    session.findById("wnd[0]").close
    WScript.Quit 11
End If

' ---------- select the IT_DATA line (F2) : try icon column, then name ----------
Dim targets
targets = Array("wnd[0]/usr/lbl[34,10]", "wnd[0]/usr/lbl[2,10]")
For try = 0 To UBound(targets)
    If Exists(targets(try)) Then
        On Error Resume Next
        session.findById(targets(try)).setFocus
        session.findById("wnd[0]").sendVKey 2
        WScript.Sleep 2500
        Err.Clear
        On Error GoTo 0
        Status "after F2 on " & targets(try) & ":"
        If session.Info.Program <> "SAPMSSY0" Or session.Info.ScreenNumber <> 120 Or Exists("wnd[1]") Then Exit For
    End If
Next

If Exists("wnd[1]") Then
    log.WriteLine "---- POPUP wnd[1] [" & session.findById("wnd[1]").Text & "]"
    DumpTree session.findById("wnd[1]"), 6, 0
    session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\SE37_it_data_editor_popup.png", 2
    session.findById("wnd[1]").sendVKey 12
    WScript.Sleep 1000
Else
    log.WriteLine "---- tbar[1]"
    Tbar1
    log.WriteLine "---- usr"
    DumpTree session.findById("wnd[0]/usr"), 7, 0
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\SE37_it_data_editor.png", 2
End If

' ---------- leave without executing ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
If Exists("wnd[1]") Then
    log.WriteLine "popup on exit [" & session.findById("wnd[1]").Text & "] - cancel"
    session.findById("wnd[1]").sendVKey 12
    WScript.Sleep 800
End If
Err.Clear
On Error GoTo 0
Status "end:"
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_se37_table done"
WScript.Quit 0
