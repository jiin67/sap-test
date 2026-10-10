' inspect_textsym2.vbs : 텍스트기호 전체 읽기 (스크롤바 사용, 읽기전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i, r, tbl, pfx, pos, k, t, lastk
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_textsym2.txt", True, True)
If connection.Children.Count > 1 Then
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
Else
    n0 = connection.Children.Count
    base.createSession
    For i = 1 To 40
        WScript.Sleep 500
        If connection.Children.Count > n0 Then Exit For
    Next
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    WScript.Sleep 1500
End If
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z33SDR600"
session.findById("wnd[0]/usr/radRS38M-FUNC_TEXT").Select
session.findById("wnd[0]/usr/btnSHOP").press
WScript.Sleep 2000
pfx = "wnd[0]/usr/tabsTX_TABSTR_CONTROL/tabpIIII/ssubSCREEN_HEADER:SAPLSETXP:1210/"
lastk = ""
For pos = 0 To 60 Step 20
    Set tbl = session.findById(pfx & "tblSAPLSETXPTSYMBOL")
    tbl.verticalScrollbar.position = pos
    WScript.Sleep 800
    On Error Resume Next
    For r = 0 To 25
        k = session.findById(pfx & "tblSAPLSETXPTSYMBOL/txtRS38M-INUM[0," & r & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        t = session.findById(pfx & "tblSAPLSETXPTSYMBOL/txtRS38M-ITEX132[1," & r & "]").Text
        If k = "" Then Exit For
        log.WriteLine k & " = [" & t & "]"
    Next
    On Error GoTo 0
    log.WriteLine "-- page pos=" & pos
Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
log.Close
session.findById("wnd[0]").close
WScript.Echo "textsym2 done"
