' inspect_zsd600_4.vbs : 버젼관리 + 화면100 ALV 구조 + 라디오별 선택화면 변화 (읽기전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zsd600_4.txt", True, True)

' 남은 세션 재사용, 없으면 생성
If connection.Children.Count > 1 Then
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    log.WriteLine "기존 추가 세션 재사용: " & session.Id
Else
    n0 = connection.Children.Count
    base.createSession
    For i = 1 To 40
        WScript.Sleep 500
        If connection.Children.Count > n0 Then Exit For
    Next
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    WScript.Sleep 1500
    log.WriteLine "새 세션: " & session.Id
End If

Sub Dump(comp, depth)
    Dim c, k, line
    line = String(depth * 2, " ") & comp.Type
    On Error Resume Next
    line = line & "/" & comp.SubType
    On Error GoTo 0
    line = line & " | " & comp.Name
    On Error Resume Next
    line = line & " | [" & comp.Text & "]"
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

Function FindMenu(comp, txt)
    Dim c, k, r
    Set FindMenu = Nothing
    On Error Resume Next
    If InStr(comp.Text, txt) > 0 Then Set FindMenu = comp : Exit Function
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Function
    On Error GoTo 0
    For c = 0 To k - 1
        Set r = FindMenu(comp.Children.ElementAt(c), txt)
        If Not r Is Nothing Then Set FindMenu = r : Exit Function
    Next
End Function

Function FindGrid(comp)
    Dim c, k, r, st
    Set FindGrid = Nothing
    On Error Resume Next
    st = "" : st = comp.SubType
    If st = "GridView" Then Set FindGrid = comp : Exit Function
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Function
    On Error GoTo 0
    For c = 0 To k - 1
        Set r = FindGrid(comp.Children.ElementAt(c))
        If Not r Is Nothing Then Set FindGrid = r : Exit Function
    Next
End Function

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub

Sub DumpPopups()
    Dim w, n
    On Error Resume Next
    For n = 1 To 3
        Set w = session.findById("wnd[" & n & "]")
        If Err.Number = 0 Then
            log.WriteLine "--- popup wnd[" & n & "] ---"
            Dump w, 0
        End If
        Err.Clear
    Next
    On Error GoTo 0
End Sub

Sub DumpGrid()
    Dim grid, cols, c, r, maxr
    Set grid = FindGrid(session.findById("wnd[0]"))
    If grid Is Nothing Then
        log.WriteLine "(GridView 없음)"
        Exit Sub
    End If
    On Error Resume Next
    log.WriteLine "--- ALV grid: " & grid.Id & "  rows=" & grid.RowCount & " cols=" & grid.ColumnCount
    Set cols = grid.ColumnOrder
    For c = 0 To cols.Count - 1
        log.WriteLine "  col " & c & ": " & cols(c) & " | " & grid.GetColumnTitles(cols(c))(0)
    Next
    maxr = grid.RowCount - 1
    If maxr > 4 Then maxr = 4
    For r = 0 To maxr
        log.WriteLine "--- row" & r & " ---"
        For c = 0 To cols.Count - 1
            log.WriteLine "  " & cols(c) & " = [" & grid.GetCellValue(r, cols(c)) & "]"
        Next
    Next
    ' 툴바 버튼
    Dim tb, b
    Set tb = grid.ToolbarButtonCount
    For b = 0 To grid.ToolbarButtonCount - 1
        log.WriteLine "  gridbtn " & b & ": id=" & grid.GetToolbarButtonId(b) & " text=[" & grid.GetToolbarButtonText(b) & "] tip=[" & grid.GetToolbarButtonTooltip(b) & "]"
    Next
    On Error GoTo 0
End Sub

' 1) 버젼관리
log.WriteLine "=============== SE38 버젼관리 ==============="
Tcode "/nSE38"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z33SDR600"
Dim m
Set m = FindMenu(session.findById("wnd[0]/mbar"), "버젼관리")
If m Is Nothing Then
    log.WriteLine "메뉴 못 찾음"
Else
    m.Select
    WScript.Sleep 2500
    Status
    DumpPopups
    Dump session.findById("wnd[0]/usr"), 0
End If

' 2) 주문조회 → 화면 100
log.WriteLine "=============== 주문조회 실행 → 결과화면 ==============="
Tcode "/nZSD600"
With session
    .findById("wnd[0]/usr/radP_RBLIST").Select
    .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
    .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.09.08"
    .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
    .findById("wnd[0]/tbar[1]/btn[8]").press
End With
WScript.Sleep 5000
Status
DumpPopups
Dump session.findById("wnd[0]/tbar[1]"), 0
Dump session.findById("wnd[0]/usr"), 0
DumpGrid

' 3) 주문로그 라디오 → 선택화면 변화
log.WriteLine "=============== 주문로그 라디오 선택 후 선택화면 ==============="
Tcode "/nZSD600"
session.findById("wnd[0]/usr/radP_RBLOG").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
Status
DumpPopups
Dump session.findById("wnd[0]/usr"), 0

log.WriteLine "=============== 주문로그 실행 ==============="
On Error Resume Next
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.09.08"
Err.Clear
On Error GoTo 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 5000
Status
DumpPopups
Dump session.findById("wnd[0]/tbar[1]"), 0
Dump session.findById("wnd[0]/usr"), 0
DumpGrid

' 4) 데이터 동기화 라디오 → 선택화면 변화 (실행 안 함)
log.WriteLine "=============== 데이터 동기화 라디오 선택 후 선택화면 (실행 안 함) ==============="
Tcode "/nZSD600"
session.findById("wnd[0]/usr/radP_RBSYNC").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
Status
DumpPopups
Dump session.findById("wnd[0]/usr"), 0

Tcode "/n"
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect4 done"
