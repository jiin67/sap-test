' Q_trigger.vbs : open ZSD600 query for one BSTKD on KDQ in a fresh session and leave it open (used to surface the debugger)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, n0, i, ci, b
b = "1757"
If WScript.Arguments.Count > 0 Then b = WScript.Arguments(0)
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For ci = 0 To application.Children.Count - 1
    If application.Children.ElementAt(ci).Children.ElementAt(0).Info.SystemName = "KDQ" Then Set connection = application.Children.ElementAt(ci) : Exit For
Next
If connection Is Nothing Then WScript.Echo "NO KDQ" : WScript.Quit 99
Set base = connection.Children.ElementAt(0)
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZSD600"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
With session
    .findById("wnd[0]/usr/radP_RBLIST").Select
    .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = b
    .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
    .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
    .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
    .findById("wnd[0]/tbar[1]/btn[8]").press
End With
WScript.Sleep 2000
WScript.Echo "screen=" & session.Info.ScreenNumber & " sbar=" & session.findById("wnd[0]/sbar").Text
