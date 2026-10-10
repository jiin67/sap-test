' close_idle_sessions.vbs : close sessions created by test scripts that sit idle on SESSION_MANAGER. Never touches ses[0].
Option Explicit
Dim SapGuiAuto, application, connection, s, i, n
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
If connection.Children.ElementAt(0).Info.SystemName <> "KDQ" Then WScript.Quit 99
n = connection.Children.Count
For i = n - 1 To 1 Step -1
    Set s = connection.Children.ElementAt(i)
    If s.Info.Transaction = "SESSION_MANAGER" And s.Info.SystemName = "KDQ" Then
        WScript.Echo "closing " & s.Id
        s.findById("wnd[0]").close
        WScript.Sleep 1500
    Else
        WScript.Echo "keep " & s.Id & " tcode=" & s.Info.Transaction
    End If
Next
WScript.Echo "sessions now=" & connection.Children.Count