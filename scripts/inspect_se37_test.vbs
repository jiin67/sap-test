' inspect_se37_test.vbs : READ-ONLY. Open the SE37 single-test screen of the RFC and dump it.
'   Does NOT execute the function module. log: logs\inspect_se37_test.txt (UTF-16)
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_se37_test.txt", True, True)

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
    s = s & " tip=[" & comp.Tooltip & "]"
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

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE37"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtRS38L-NAME").Text = "ZCRMF03_RFC_ZCRM_CHRG_CNFRM"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
Status "after F8 (test screen):"
If Exists("wnd[1]") Then
    log.WriteLine "POPUP [" & session.findById("wnd[1]").Text & "]"
    DumpTree session.findById("wnd[1]"), 4, 0
    session.findById("wnd[1]").sendVKey 12
    WScript.Sleep 1000
    Status "after popup cancel:"
End If
log.WriteLine "---- tbar[1]"
For i = 0 To 50
    If Exists("wnd[0]/tbar[1]/btn[" & i & "]") Then
        On Error Resume Next
        log.WriteLine "  btn[" & i & "] text=[" & session.findById("wnd[0]/tbar[1]/btn[" & i & "]").Text & "] tip=[" & session.findById("wnd[0]/tbar[1]/btn[" & i & "]").Tooltip & "]"
        Err.Clear
        On Error GoTo 0
    End If
Next
log.WriteLine "---- usr"
DumpTree session.findById("wnd[0]/usr"), 6, 0
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\SE37_test_screen.png", 2

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
Err.Clear
On Error GoTo 0
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_se37_test done"
WScript.Quit 0
