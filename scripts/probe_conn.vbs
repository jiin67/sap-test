' probe_conn.vbs : list connections/sessions (ASCII only output)
Option Explicit
Dim SapGuiAuto, application, connection, session, i, j, out

On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then WScript.Echo "ERR: no SAPGUI object" : WScript.Quit 1
Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Then WScript.Echo "ERR: no scripting engine" : WScript.Quit 2
On Error GoTo 0

WScript.Echo "connections=" & application.Children.Count
For i = 0 To application.Children.Count - 1
    Set connection = application.Children.ElementAt(i)
    WScript.Echo "conn[" & i & "] sessions=" & connection.Children.Count
    For j = 0 To connection.Children.Count - 1
        Set session = connection.Children.ElementAt(j)
        On Error Resume Next
        out = "  ses[" & j & "] id=" & session.Id & _
              " sys=" & session.Info.SystemName & _
              " client=" & session.Info.Client & _
              " user=" & session.Info.User & _
              " lang=" & session.Info.Language & _
              " tcode=" & session.Info.Transaction & _
              " prog=" & session.Info.Program & _
              " screen=" & session.Info.ScreenNumber
        If Err.Number <> 0 Then out = "  ses[" & j & "] ERR " & Err.Description : Err.Clear
        On Error GoTo 0
        WScript.Echo out
    Next
Next
WScript.Quit 0
