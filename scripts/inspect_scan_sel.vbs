' inspect_scan_sel.vbs : RPR_ABAP_SOURCE_SCAN 선택화면 덤프 (읽기전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_scan_sel.txt", True, True)
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
Sub Dump(comp, depth)
    Dim c, k, line
    If comp.Type = "GuiMenubar" Then Exit Sub
    line = String(depth * 2, " ") & comp.Type & " | " & comp.Name
    On Error Resume Next
    line = line & " | [" & Left(comp.Text, 60) & "]"
    On Error GoTo 0
    log.WriteLine line
    On Error Resume Next
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k - 1
        Dump comp.Children.ElementAt(c), depth + 1
    Next
End Sub
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "RPR_ABAP_SOURCE_SCAN"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 2000
log.WriteLine "screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
Dump session.findById("wnd[0]/usr"), 0
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
log.Close
session.findById("wnd[0]").close
WScript.Echo "done"
