' S48_capture_dump.vbs : on the last KDQ session, HardCopy the current (dump) screen, dump label text, then exit to menu. read-only-ish (only /n).
Option Explicit
Dim SapGuiAuto, application, connection, session, ci, c, comp, usr, fso, log
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For ci = 0 To application.Children.Count - 1
    If application.Children.ElementAt(ci).Children.ElementAt(0).Info.SystemName = "KDQ" Then Set connection = application.Children.ElementAt(ci) : Exit For
Next
If connection Is Nothing Then WScript.Echo "NO KDQ" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S48_dump.txt", True, True)
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
log.WriteLine "session=" & session.Id & " prog=" & session.Info.Program & " screen=" & session.Info.ScreenNumber & " tcode=" & session.Info.Transaction
On Error Resume Next
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\gi\DELETE_02_1722_ABAP_DUMP.png", 2
Err.Clear
On Error GoTo 0
Set usr = session.findById("wnd[0]/usr")
Dim n : n = 0
For c = 0 To usr.Children.Count - 1
    Set comp = usr.Children.ElementAt(c)
    On Error Resume Next
    If comp.Type = "GuiLabel" And Trim(comp.Text) <> "" Then
        log.WriteLine comp.Text
        n = n + 1
        If n > 120 Then Exit For
    End If
    On Error GoTo 0
Next
log.WriteLine "--- labels=" & n
' leave the dump screen
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800
log.WriteLine "after /n screen=" & session.Info.Program & "/" & session.Info.ScreenNumber
log.Close
WScript.Echo "dump captured"
