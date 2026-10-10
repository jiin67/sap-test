' inspect_zsd600_5.vbs : 버젼관리 그리드 내용 + 화면100 전체 트리/ALV 내용 (읽기전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zsd600_5.txt", True, True)

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
log.WriteLine "세션: " & session.Id

Sub Dump(comp, depth)
    Dim c, k, line
    If comp.Type = "GuiMenubar" Then Exit Sub
    line = String(depth * 2, " ") & comp.Type
    On Error Resume Next
    line = line & "/" & comp.SubType
    On Error GoTo 0
    line = line & " | " & comp.Id
    On Error Resume Next
    line = line & " | [" & Left(comp.Text, 80) & "]"
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

Sub DumpGridAt(id, maxrows)
    Dim grid, cols, c, r, maxr, b
    On Error Resume Next
    Set grid = session.findById(id)
    If Err.Number <> 0 Then log.WriteLine "(grid 없음: " & id & ")" : Err.Clear : Exit Sub
    log.WriteLine "--- ALV grid: " & id & "  rows=" & grid.RowCount & " cols=" & grid.ColumnCount
    Set cols = grid.ColumnOrder
    For c = 0 To cols.Count - 1
        log.WriteLine "  col " & c & ": " & cols(c) & " | " & grid.GetColumnTitles(cols(c))(0)
    Next
    maxr = grid.RowCount - 1
    If maxr > maxrows - 1 Then maxr = maxrows - 1
    For r = 0 To maxr
        Dim line : line = "  row" & r & ":"
        For c = 0 To cols.Count - 1
            line = line & " " & cols(c) & "=[" & grid.GetCellValue(r, cols(c)) & "]"
        Next
        log.WriteLine line
    Next
    For b = 0 To grid.ToolbarButtonCount - 1
        log.WriteLine "  gridbtn " & b & ": id=" & grid.GetToolbarButtonId(b) & " text=[" & grid.GetToolbarButtonText(b) & "] tip=[" & grid.GetToolbarButtonTooltip(b) & "]"
    Next
    On Error GoTo 0
End Sub

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub

' 1) 버젼관리 그리드
log.WriteLine "=============== SE38 버젼관리 그리드 ==============="
Tcode "/nSE38"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z33SDR600"
Dim m
Set m = FindMenu(session.findById("wnd[0]/mbar"), "버젼관리")
m.Select
WScript.Sleep 2500
Status
DumpGridAt "wnd[0]/usr/cntlCONT/shellcont/shell/shellcont[0]/shell", 30

' 2) 주문조회 → 화면 100 전체 트리
log.WriteLine "=============== 주문조회 결과화면 wnd[0] 전체 트리 ==============="
Tcode "/nZSD600"
With session
    .findById("wnd[0]/usr/radP_RBLIST").Select
    .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
    .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.09.08"
    .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
    .findById("wnd[0]/tbar[1]/btn[8]").press
End With
WScript.Sleep 6000
Status
Dump session.findById("wnd[0]"), 0
' 흔한 도킹 컨테이너 경로 시도
DumpGridAt "wnd[0]/shellcont/shell", 5
DumpGridAt "wnd[0]/shellcont[0]/shell", 5
DumpGridAt "wnd[0]/shellcont/shell/shellcont[0]/shell", 5
DumpGridAt "wnd[0]/shellcont/shell/shellcont[1]/shell", 5

' 3) 주문로그 → 화면 100
log.WriteLine "=============== 주문로그 결과화면 wnd[0] 전체 트리 ==============="
Tcode "/nZSD600"
session.findById("wnd[0]/usr/radP_RBLOG").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.09.08"
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 6000
Status
Dump session.findById("wnd[0]"), 0
DumpGridAt "wnd[0]/shellcont/shell", 5
DumpGridAt "wnd[0]/shellcont[0]/shell", 5
DumpGridAt "wnd[0]/shellcont/shell/shellcont[0]/shell", 5
DumpGridAt "wnd[0]/shellcont/shell/shellcont[1]/shell", 5

Tcode "/n"
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect5 done"
