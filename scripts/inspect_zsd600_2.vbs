' inspect_zsd600_2.vbs : SE38 속성 / 버전이력 / 주문조회(읽기전용) 결과 구조 덤프
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zsd600_2.txt", True, True)

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500

Sub Dump(comp, depth)
    Dim c, k, line, chg
    line = String(depth * 2, " ") & comp.Type & " | " & comp.Name
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

' 메뉴에서 텍스트로 항목 찾기
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

' 1) SE38 속성
log.WriteLine "=============== SE38 속성 ==============="
Tcode "/nSE38"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z33SDR600"
session.findById("wnd[0]/usr/radRS38M-FUNC_HEAD").Select
session.findById("wnd[0]/usr/btnSHOP").press
WScript.Sleep 1500
Status
DumpPopups
Dump session.findById("wnd[0]/usr"), 0

' 2) 버전 관리
log.WriteLine "=============== 버전 관리 ==============="
Dim m
Set m = FindMenu(session.findById("wnd[0]/mbar"), "버전 관리")
If m Is Nothing Then
    log.WriteLine "메뉴 '버전 관리' 못 찾음. 메뉴 트리:"
    Dump session.findById("wnd[0]/mbar"), 0
Else
    log.WriteLine "메뉴 선택: " & m.Id
    m.Select
    WScript.Sleep 1500
    Status
    DumpPopups
    Dump session.findById("wnd[0]/usr"), 0
End If

' 3) ZSD600 주문조회 (읽기전용) 기본값 실행
log.WriteLine "=============== ZSD600 주문조회 기본값 실행 ==============="
Tcode "/nZSD600"
session.findById("wnd[0]/usr/radP_RBLIST").Select
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 4000
Status
DumpPopups
Dump session.findById("wnd[0]/tbar[1]"), 0
Dump session.findById("wnd[0]/usr"), 0
' ALV 그리드가 있으면 컬럼/행수 출력
Dim grid, cols, c
On Error Resume Next
Set grid = Nothing
Set grid = FindGrid(session.findById("wnd[0]/usr"))
If Not grid Is Nothing Then
    log.WriteLine "--- ALV grid: " & grid.Id & "  rows=" & grid.RowCount & " cols=" & grid.ColumnCount
    Set cols = grid.ColumnOrder
    For c = 0 To cols.Count - 1
        log.WriteLine "  col " & c & ": " & cols(c) & " | " & grid.GetColumnTitles(cols(c))(0)
    Next
    If grid.RowCount > 0 Then
        log.WriteLine "--- row0 ---"
        For c = 0 To cols.Count - 1
            log.WriteLine "  " & cols(c) & " = [" & grid.GetCellValue(0, cols(c)) & "]"
        Next
    End If
End If
Err.Clear
On Error GoTo 0

Function FindGrid(comp)
    Dim c, k, r
    Set FindGrid = Nothing
    On Error Resume Next
    If comp.Type = "GuiGridView" Then Set FindGrid = comp : Exit Function
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Function
    On Error GoTo 0
    For c = 0 To k - 1
        Set r = FindGrid(comp.Children.ElementAt(c))
        If Not r Is Nothing Then Set FindGrid = r : Exit Function
    Next
End Function

Tcode "/n"
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect2 done"
