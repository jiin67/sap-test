' inspect_src.vbs : SE38 소스/텍스트요소 읽기 시도 (읽기전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_src.txt", True, True)
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

Sub Dump(comp, depth)
    Dim c, k, line
    If comp.Type = "GuiMenubar" Then Exit Sub
    line = String(depth * 2, " ") & comp.Type
    On Error Resume Next
    line = line & "/" & comp.SubType
    On Error GoTo 0
    line = line & " | " & comp.Id
    On Error Resume Next
    line = line & " | [" & Left(comp.Text, 60) & "]"
    On Error GoTo 0
    log.WriteLine line
    On Error Resume Next
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k - 1
        Dump comp.Children.ElementAt(c), depth + 1
    Next
End Sub

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub

Sub Status()
    log.WriteLine "  screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & "  sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
End Sub

' 1) 소스코드 조회
log.WriteLine "=============== SE38 소스코드 조회 ==============="
Tcode "/nSE38"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z33SDR600"
session.findById("wnd[0]/usr/radRS38M-FUNC_EDIT").Select
session.findById("wnd[0]/usr/btnSHOP").press
WScript.Sleep 2500
Status
Dump session.findById("wnd[0]"), 0
' 에디터 텍스트 읽기 시도
Dim ids, id, ed, txt
ids = Array("wnd[0]/shellcont/shell", "wnd[0]/shellcont[0]/shell", "wnd[0]/shellcont[1]/shell", "wnd[0]/usr/cntlEDITOR/shellcont/shell", "wnd[0]/usr/cntlEDITOR/shellcont/shell/shellcont[1]/shell")
For Each id In ids
    On Error Resume Next
    Set ed = session.findById(id)
    If Err.Number = 0 Then
        Err.Clear
        txt = ed.Text
        If Err.Number = 0 Then
            log.WriteLine "--- " & id & " .Text len=" & Len(txt)
            log.WriteLine Left(txt, 200)
        Else
            log.WriteLine "--- " & id & " .Text 실패: " & Err.Description
        End If
        Err.Clear
        log.WriteLine "    LineCount=" & ed.LineCount
        Err.Clear
    End If
    Err.Clear
    On Error GoTo 0
Next

' 2) 텍스트요소 조회
log.WriteLine "=============== SE38 텍스트요소 조회 ==============="
Tcode "/nSE38"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z33SDR600"
session.findById("wnd[0]/usr/radRS38M-FUNC_TEXT").Select
session.findById("wnd[0]/usr/btnSHOP").press
WScript.Sleep 2500
Status
Dump session.findById("wnd[0]"), 0

Tcode "/n"
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_src done"
