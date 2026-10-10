' inspect_fbgendat.vbs : READ-ONLY. Open SE38 -> FBGENDAT selection screen in KDD and dump it. No execute.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For i = 0 To application.Children.Count - 1
    Set base = application.Children.ElementAt(i).Children.ElementAt(0)
    If base.Info.SystemName = "KDD" Then Set connection = application.Children.ElementAt(i) : Exit For
Next
If connection Is Nothing Then WScript.Echo "GUARD no KDD" : WScript.Quit 99
Set base = connection.Children.ElementAt(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_fbgendat.txt", True, True)
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDD" Then WScript.Quit 99
log.WriteLine "sys=" & session.Info.SystemName & " client=" & session.Info.Client & " user=" & session.Info.User & " session=" & session.Id

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
    s = s & " tip=[" & comp.Tooltip & "] chg=" & comp.Changeable
    log.WriteLine s
    Err.Clear
    If lvl >= depth Then Exit Sub
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k - 1
        DumpTree comp.Children.ElementAt(c), depth, lvl + 1
    Next
End Sub

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
Err.Clear
On Error GoTo 0
Status "SE38:"
DumpTree session.findById("wnd[0]/usr"), 3, 0
On Error Resume Next
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "FBGENDAT"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 2500
Err.Clear
On Error GoTo 0
Status "after F8:"
If Exists("wnd[1]") Then
    log.WriteLine "POPUP [" & session.findById("wnd[1]").Text & "]"
    DumpTree session.findById("wnd[1]"), 4, 0
End If
DumpTree session.findById("wnd[0]/usr"), 4, 0
log.WriteLine "---- tbar[1]"
DumpTree session.findById("wnd[0]/tbar[1]"), 2, 0
log.WriteLine "---- mbar"
DumpTree session.findById("wnd[0]/mbar"), 2, 0
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\FBG_probe_sel.png", 2
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
log.Close
session.findById("wnd[0]").close
WScript.Echo "done"
