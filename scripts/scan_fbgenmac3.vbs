' scan_fbgenmac3.vbs : READ-ONLY. verify scanner + search FBGENMAC in ZCRM* function groups (KDD).
Option Explicit
Const LOGNAME = "scan_fbgenmac3"
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
Sub Scan(str, fg)
    Tcode "/nSE38"
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "RS_ABAP_SOURCE_SCAN"
    session.findById("wnd[0]").sendVKey 8
    WScript.Sleep 2000
    session.findById("wnd[0]/usr/txtSSTRING-LOW").Text = str
    session.findById("wnd[0]/usr/ctxtFUNCGRP-LOW").Text = fg
    session.findById("wnd[0]/usr/chkINCLU").Selected = True
    session.findById("wnd[0]/usr/chkCOMMENT").Selected = False
    session.findById("wnd[0]").sendVKey 8
    WScript.Sleep 8000
    Status "scan [" & str & "] in [" & fg & "]:"
    If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : DumpTree session.findById("wnd[1]"), 4, 0 : session.findById("wnd[1]").sendVKey 0 : WScript.Sleep 1000
    Dim k, u
    Set u = session.findById("wnd[0]/usr")
    If session.Info.Program <> "RS_ABAP_SOURCE_SCAN" Or session.Info.ScreenNumber <> 1000 Then
        For k = 0 To u.Children.Count - 1
            If Trim(u.Children.ElementAt(k).Text) <> "" Then log.WriteLine "  " & Mid(u.Children.ElementAt(k).Id, InStrRev(u.Children.ElementAt(k).Id, "/") + 1) & " [" & u.Children.ElementAt(k).Text & "]"
        Next
    End If
End Sub
Scan "CHRG_CNFRM", "ZCRMG03"
Scan "FBGENMAC", "ZCRM*"
Scan "FBGENMAC", "Z33*"
Finish 0
