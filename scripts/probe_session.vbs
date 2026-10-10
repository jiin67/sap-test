' probe_session.vbs : call createSession on KDQ base and watch all connections/sessions for 20s (read only)
Option Explicit
Dim SapGuiAuto, application, connection, base, ci, si, t, s, c
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Sub ListAll(tag)
    Dim i, j, cn, sn
    WScript.Echo tag & " connections=" & application.Children.Count
    For i = 0 To application.Children.Count - 1
        Set cn = application.Children.ElementAt(i)
        WScript.Echo "  con" & i & " [" & cn.Description & "] sessions=" & cn.Children.Count
        For j = 0 To cn.Children.Count - 1
            Set sn = cn.Children.ElementAt(j)
            On Error Resume Next
            WScript.Echo "    " & sn.Id & " sys=" & sn.Info.SystemName & " client=" & sn.Info.Client & " user=" & sn.Info.User & " tcode=" & sn.Info.Transaction & " screen=" & sn.Info.ScreenNumber & " w1=" & sn.findById("wnd[1]").Text
            Err.Clear
            On Error GoTo 0
        Next
    Next
End Sub
ListAll "BEFORE"
Set connection = Nothing
For ci = 0 To application.Children.Count - 1
    If application.Children.ElementAt(ci).Children.ElementAt(0).Info.SystemName = "KDQ" Then Set connection = application.Children.ElementAt(ci) : Exit For
Next
If connection Is Nothing Then WScript.Echo "NO KDQ" : WScript.Quit 99
Set base = connection.Children.ElementAt(0)
WScript.Echo "base=" & base.Id & " calling createSession..."
On Error Resume Next
base.createSession
WScript.Echo "createSession err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0
For t = 1 To 10
    WScript.Sleep 2000
    ListAll "T+" & (t*2) & "s"
Next
