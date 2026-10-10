' shotprobe2.vbs : READ-ONLY. Why does HardCopy fail on a session we created?
Option Explicit
Dim SapGuiAuto, application, connection, base, session, i, n0, p

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For i = 0 To application.Children.Count - 1
    If application.Children.ElementAt(i).Children.ElementAt(0).Info.SystemName = "KDQ" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then WScript.Quit 99
Set base = connection.Children.ElementAt(0)
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 2000
WScript.Echo "new session=" & session.Id

On Error Resume Next
WScript.Echo "1 plain:"
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\_p2a.png", 2
WScript.Echo "   err=" & Err.Number
Err.Clear

session.findById("wnd[0]").maximize
WScript.Sleep 1000
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\_p2b.png", 2
WScript.Echo "2 maximize err=" & Err.Number
Err.Clear

application.ActiveSession = session
WScript.Sleep 1000
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\_p2c.png", 2
WScript.Echo "3 activesession err=" & Err.Number
Err.Clear

session.findById("wnd[0]/usr").SetFocus
WScript.Sleep 800
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\_p2d.png", 2
WScript.Echo "4 setfocus err=" & Err.Number
Err.Clear

WScript.Echo "busy=" & session.Busy & " visible? id=" & session.Id
session.findById("wnd[0]").close
Err.Clear
On Error GoTo 0
