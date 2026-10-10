' inspect_zcrm052_01.vbs : existence + selection screen structure of ZCRM052 / ZCRMR052S
' READ ONLY. Opens its own session, closes it at the end.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then
    WScript.Echo "GUARD: system " & base.Info.SystemName & " not allowed"
    WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zcrm052_01.txt", True, True)
log.WriteLine "system=" & base.Info.SystemName & " client=" & base.Info.Client & " user=" & base.Info.User

' ---- open a new session so the user's screen is untouched ----
n0 = connection.Children.Count
' Always work in a session we opened ourselves. Re-using a leftover session can
' block for a long time when that session still holds a huge ALV result.
If False Then
    Set session = connection.Children.ElementAt(n0 - 1)
Else
    base.createSession
    For i = 1 To 40
        WScript.Sleep 500
        If connection.Children.Count > n0 Then Exit For
    Next
    If connection.Children.Count <= n0 Then
        log.WriteLine "FAIL: cannot create new session"
        log.Close
        WScript.Quit 10
    End If
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    WScript.Sleep 1500
End If
log.WriteLine "work session=" & session.Id

Sub Status()
    log.WriteLine "  prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & "  tcode=" & session.Info.Transaction
    log.WriteLine "  title=[" & session.findById("wnd[0]").Text & "]"
    log.WriteLine "  sbar =[" & session.findById("wnd[0]/sbar").Text & "] type=" & session.findById("wnd[0]/sbar").MessageType
End Sub

Sub Dump(comp, depth)
    Dim c, k, line, chg
    On Error Resume Next
    line = String(depth * 2, " ") & comp.Type & " | " & comp.Id
    line = line & " | name=" & comp.Name
    line = line & " | text=[" & comp.Text & "]"
    chg = comp.Changeable
    If Err.Number = 0 Then line = line & " | chg=" & chg
    Err.Clear
    log.WriteLine line
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k - 1
        Dump comp.Children.ElementAt(c), depth + 1
    Next
End Sub

Sub CloseExtraPopups()
    Dim w, guard
    guard = 0
    Do
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number <> 0 Then Err.Clear : Exit Do
        On Error GoTo 0
        log.WriteLine "  [popup] title=[" & w.Text & "]"
        Dump w, 1
        w.close
        WScript.Sleep 600
        guard = guard + 1
    Loop While guard < 3
End Sub

' ---- 1. SE38 attributes of ZCRMR052 ----
Sub ProgAttr(pname)
    log.WriteLine "=============== SE38 ATTR " & pname & " ==============="
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = pname
    session.findById("wnd[0]/usr/radRS38M-FUNC_ATTR").Select
    session.findById("wnd[0]/tbar[1]/btn[7]").press
    WScript.Sleep 1800
    If Err.Number <> 0 Then log.WriteLine "  ERR " & Err.Number & " " & Err.Description : Err.Clear
    On Error GoTo 0
    Status
    CloseExtraPopups
    On Error Resume Next
    Dump session.findById("wnd[0]/usr"), 0
    Err.Clear
    On Error GoTo 0
End Sub

ProgAttr "ZCRMR052"
ProgAttr "ZCRMR052S"

' ---- 2. SE93 for tcode ZCRM052 ----
Sub TcodeAttr(tname)
    log.WriteLine "=============== SE93 " & tname & " ==============="
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE93"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtTSTC-TCODE").Text = tname
    session.findById("wnd[0]/tbar[1]/btn[7]").press
    WScript.Sleep 1800
    If Err.Number <> 0 Then log.WriteLine "  ERR " & Err.Number & " " & Err.Description : Err.Clear
    On Error GoTo 0
    Status
    On Error Resume Next
    Dump session.findById("wnd[0]/usr"), 0
    Err.Clear
    On Error GoTo 0
    CloseExtraPopups
End Sub

TcodeAttr "ZCRM052"

' ---- 3. selection screen of ZCRM052 ----
log.WriteLine "=============== /nZCRM052 SELECTION SCREEN ==============="
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2500
If Err.Number <> 0 Then log.WriteLine "  ERR " & Err.Number & " " & Err.Description : Err.Clear
On Error GoTo 0
Status
CloseExtraPopups
On Error Resume Next
Dump session.findById("wnd[0]"), 0
Err.Clear
On Error GoTo 0

' ---- reset and close ----
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 900
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_zcrm052_01 done"
