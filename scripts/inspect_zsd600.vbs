' inspect_zsd600.vbs : 새 세션을 열어 ZSD600 선택화면 구조 + SE38 속성을 덤프 (읽기 전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i, t

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then
    WScript.Echo "GUARD: system " & base.Info.SystemName & " not allowed"
    WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zsd600.txt", True, True)

' 새 세션 생성
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
If connection.Children.Count <= n0 Then
    log.WriteLine "새 세션 생성 실패 (세션 수 제한?)"
    log.Close
    WScript.Quit 10
End If
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
log.WriteLine "새 세션: " & session.Id

Sub Dump(comp, depth)
    Dim c, k, line, chg
    line = String(depth * 2, " ") & comp.Type & " | " & comp.Id
    On Error Resume Next
    line = line & " | name=" & comp.Name
    line = line & " | text=[" & comp.Text & "]"
    chg = comp.Changeable
    If Err.Number = 0 Then line = line & " | changeable=" & chg
    Err.Clear
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

Sub Status()
    log.WriteLine "  screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & "  tcode=" & session.Info.Transaction
    log.WriteLine "  title=[" & session.findById("wnd[0]").Text & "]"
    log.WriteLine "  sbar =[" & session.findById("wnd[0]/sbar").Text & "]  type=" & session.findById("wnd[0]/sbar").MessageType
End Sub

' 1) ZSD600 선택화면
log.WriteLine "=============== /nZSD600 ==============="
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZSD600"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
Status
Dump session.findById("wnd[0]"), 0
On Error Resume Next
Dim w1 : Set w1 = session.findById("wnd[1]")
If Err.Number = 0 Then
    log.WriteLine "--- 팝업 wnd[1] ---"
    Dump w1, 0
End If
Err.Clear
On Error GoTo 0

' 2) SE38 속성 화면
log.WriteLine "=============== /nSE38 Z33SDR600 속성 ==============="
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
On Error Resume Next
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z33SDR600"
session.findById("wnd[0]/usr/radRS38M-FUNC_ATTR").Select
session.findById("wnd[0]/tbar[1]/btn[7]").press   ' Display
WScript.Sleep 1500
Status
Dump session.findById("wnd[0]/usr"), 0
Err.Clear
On Error GoTo 0

' 3) 텍스트 엘리먼트(선택 텍스트) 화면
log.WriteLine "=============== SE38 텍스트 엘리먼트 ==============="
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
On Error Resume Next
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z33SDR600"
session.findById("wnd[0]/usr/radRS38M-FUNC_TEXT").Select
session.findById("wnd[0]/tbar[1]/btn[7]").press
WScript.Sleep 1500
Status
Dump session.findById("wnd[0]/usr"), 0
Err.Clear
On Error GoTo 0

' 세션 초기화 후 종료
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect done"
