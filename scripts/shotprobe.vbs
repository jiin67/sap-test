' shotprobe.vbs : READ-ONLY. Diagnose why GuiFrameWindow.HardCopy stopped writing files.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, i, p, r

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
Set session = connection.Children.ElementAt(0)
WScript.Echo "sessions=" & connection.Children.Count & " sysname=" & session.Info.SystemName

p = "C:\sap-test\results\shots\_probe1.png"
On Error Resume Next
r = session.findById("wnd[0]").HardCopy(p, 2)
WScript.Echo "A err=" & Err.Number & " [" & Err.Description & "] ret=[" & r & "]"
Err.Clear

p = "C:\Users\jiink\AppData\Local\Temp\_probe2.png"
r = session.findById("wnd[0]").HardCopy(p, 2)
WScript.Echo "B err=" & Err.Number & " [" & Err.Description & "] ret=[" & r & "]"
Err.Clear

session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\_probe3.png", 2
WScript.Echo "C err=" & Err.Number & " [" & Err.Description & "]"
Err.Clear
On Error GoTo 0
