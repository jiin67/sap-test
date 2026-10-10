' cleanup_kdq.vbs : on the KDQ connection only, answer No/Cancel on open popups and close extra sessions (never ses[0])
Option Explicit
Dim SapGuiAuto, application, connection, s, i, ci, guard, n
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For ci = 0 To application.Children.Count - 1
    If application.Children.ElementAt(ci).Children.ElementAt(0).Info.SystemName = "KDQ" Then Set connection = application.Children.ElementAt(ci) : Exit For
Next
If connection Is Nothing Then WScript.Echo "NO KDQ" : WScript.Quit 99
Function HasW(sess, k)
    On Error Resume Next
    HasW = False
    Dim w : Set w = sess.findById("wnd[" & k & "]")
    If Err.Number = 0 Then HasW = True
    Err.Clear
    On Error GoTo 0
End Function
Sub SafeClose(sess, wname)
    Dim usr, c, comp, pressed
    On Error Resume Next
    pressed = False
    Set usr = sess.findById(wname & "/usr")
    If Err.Number = 0 Then
        For c = 0 To usr.Children.Count - 1
            Set comp = usr.Children.ElementAt(c)
            If comp.Type = "GuiButton" Then
                If InStr(comp.Text & comp.Tooltip, ChrW(&HC544)&ChrW(&HB2C8)&ChrW(&HC624)) > 0 Or InStr(comp.Text & comp.Tooltip, ChrW(&HCDE8)&ChrW(&HC18C)) > 0 Then comp.press : pressed = True : Exit For
            End If
        Next
    End If
    Err.Clear
    If Not pressed Then sess.findById(wname & "/tbar[0]/btn[12]").press : Err.Clear
    If HasW(sess, Mid(wname, 5, 1)) Then sess.findById(wname).close : Err.Clear
    On Error GoTo 0
    WScript.Sleep 500
End Sub
For i = connection.Children.Count - 1 To 1 Step -1
    Set s = connection.Children.ElementAt(i)
    guard = 0
    Do While HasW(s, 1) And guard < 6
        SafeClose s, "wnd[3]" : SafeClose s, "wnd[2]" : SafeClose s, "wnd[1]"
        guard = guard + 1
    Loop
    On Error Resume Next
    s.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    s.findById("wnd[0]").sendVKey 0
    WScript.Sleep 500
    s.findById("wnd[0]").close
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 500
Next
WScript.Echo "kdq sessions=" & connection.Children.Count
