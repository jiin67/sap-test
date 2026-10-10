' inspect_su3b.vbs : READ-ONLY. SU3 parameter tab structure (KDD).
Option Explicit
Const LOGNAME = "inspect_su3b"
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
Set log = fso.CreateTextFile("C:\sap-test\logs\" & LOGNAME & ".txt", True, True)
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDD" Then WScript.Quit 99
log.WriteLine "sys=" & session.Info.SystemName & " client=" & session.Info.Client & " user=" & session.Info.User & " session=" & session.Id & " at=" & Now

Function Exists(id)
    Dim o
    On Error Resume Next
    Set o = session.findById(id)
    Exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
Function T(id)
    On Error Resume Next
    T = ""
    T = session.findById(id).Text
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
Sub Shot(name)
    On Error Resume Next
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\" & name & ".png", 2
    Err.Clear
    On Error GoTo 0
End Sub
Sub Tcode(t)
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1800
    Err.Clear
    On Error GoTo 0
End Sub
Sub Finish(code)
    On Error Resume Next
    Tcode "/n"
    If Exists("wnd[1]") Then log.WriteLine "popup on exit [" & T("wnd[1]") & "]" : session.findById("wnd[1]").sendVKey 12 : WScript.Sleep 800
    log.WriteLine "exit=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Echo LOGNAME & " exit=" & code
    WScript.Quit code
End Sub
Tcode "/nSU3"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpPARAM").Select
WScript.Sleep 1500
Status "param tab:"
DumpTree session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpPARAM"), 6, 0
log.WriteLine "---- tbar[0]"
DumpTree session.findById("wnd[0]/tbar[0]"), 2, 0
Shot "SU3_probe_param"
Finish 0
