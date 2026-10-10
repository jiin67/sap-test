' probe_screen.vbs : dump current screen of the LAST KDQ session (read only)
Option Explicit
Dim SapGuiAuto, application, connection, session, ci, c, comp, usr, n
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For ci = 0 To application.Children.Count - 1
    If application.Children.ElementAt(ci).Children.ElementAt(0).Info.SystemName = "KDQ" Then Set connection = application.Children.ElementAt(ci) : Exit For
Next
If connection Is Nothing Then WScript.Echo "NO KDQ" : WScript.Quit 99
WScript.Echo "sessions=" & connection.Children.Count
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Echo "id=" & session.Id & " tcode=" & session.Info.Transaction & " prog=" & session.Info.Program & " screen=" & session.Info.ScreenNumber
WScript.Echo "title=" & session.findById("wnd[0]").Text
WScript.Echo "sbar=" & session.findById("wnd[0]/sbar").Text
For n = 1 To 3
    On Error Resume Next
    WScript.Echo "wnd[" & n & "]=" & session.findById("wnd[" & n & "]").Text
    On Error GoTo 0
Next
Set usr = session.findById("wnd[0]/usr")
For c = 0 To usr.Children.Count - 1
    Set comp = usr.Children.ElementAt(c)
    If InStr(comp.Name, "%_") = 0 Then
        On Error Resume Next
        WScript.Echo "  " & comp.Type & " | " & comp.Name & " | " & comp.Text
        On Error GoTo 0
    End If
Next
