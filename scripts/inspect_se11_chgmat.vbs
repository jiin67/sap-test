' inspect_se11_chgmat.vbs : SE11 필드(F7) + 자재변경 팝업 구조(저장 없이 취소) (읽기전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_se11_chgmat.txt", True, True)
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

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Sub Status()
    log.WriteLine "  screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " title=[" & session.findById("wnd[0]").Text & "] sbar=[" & session.findById("wnd[0]/sbar").Text & "] type=" & session.findById("wnd[0]/sbar").MessageType
End Sub
Sub Dump(comp, depth)
    Dim c, k, line
    If comp.Type = "GuiMenubar" Then Exit Sub
    line = String(depth * 2, " ") & comp.Type
    On Error Resume Next
    line = line & "/" & comp.SubType
    On Error GoTo 0
    line = line & " | " & comp.Id
    On Error Resume Next
    line = line & " | [" & Left(comp.Text, 70) & "]"
    line = line & " chg=" & comp.Changeable
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
Function FindType(comp, tp)
    Dim c, k, r
    Set FindType = Nothing
    On Error Resume Next
    If comp.Type = tp Then Set FindType = comp : Exit Function
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Function
    On Error GoTo 0
    For c = 0 To k - 1
        Set r = FindType(comp.Children.ElementAt(c), tp)
        If Not r Is Nothing Then Set FindType = r : Exit Function
    Next
End Function
Sub TableDump(tblId, maxrows)
    Dim pos, r, c, line, cell, visible, lastFirst, first, tbl
    lastFirst = "###"
    Set tbl = session.findById(tblId)
    visible = tbl.VisibleRowCount
    For pos = 0 To maxrows Step visible
        On Error Resume Next
        Set tbl = session.findById(tblId)
        tbl.verticalScrollbar.position = pos
        WScript.Sleep 400
        Set tbl = session.findById(tblId)
        first = ""
        For r = 0 To visible - 1
            line = ""
            For c = 0 To tbl.Columns.Count - 1
                Set cell = tbl.GetCell(r, c)
                If Err.Number <> 0 Then Err.Clear : Exit For
                If cell.Text <> "" Then line = line & " | " & cell.Text
            Next
            If r = 0 Then first = line
            If line <> "" Then log.WriteLine "  " & line
        Next
        On Error GoTo 0
        If first = lastFirst Or first = "" Then Exit For
        lastFirst = first
    Next
End Sub
Sub GridAll(grid, maxrows)
    Dim cols, c, r, line, maxr
    On Error Resume Next
    log.WriteLine "  grid rows=" & grid.RowCount & " cols=" & grid.ColumnCount
    Set cols = grid.ColumnOrder
    maxr = grid.RowCount - 1
    If maxr > maxrows - 1 Then maxr = maxrows - 1
    For r = 0 To maxr
        line = "  r" & r & ":"
        For c = 0 To cols.Count - 1
            line = line & " " & cols(c) & "=[" & grid.GetCellValue(r, cols(c)) & "]"
        Next
        log.WriteLine line
    Next
    On Error GoTo 0
End Sub

' ---------- 1) SE11 (F7 조회) ----------
Dim tabs, tb, tbl
tabs = Array()
For Each tb In tabs
    log.WriteLine "=============== SE11: " & tb & " ==============="
    Tcode "/nSE11"
    On Error Resume Next
    session.findById("wnd[0]/usr/radRSRD1-TBMA").Select
    session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").Text = tb
    session.findById("wnd[0]").sendVKey 7
    WScript.Sleep 2500
    Status
    Dim w1 : Set w1 = session.findById("wnd[1]")
    If Err.Number = 0 Then log.WriteLine "  popup=[" & w1.Text & "]" : w1.close : WScript.Sleep 500
    Err.Clear
    Set tbl = FindType(session.findById("wnd[0]/usr"), "GuiTableControl")
    If tbl Is Nothing Then
        log.WriteLine "  테이블컨트롤 없음 - usr 덤프"
        Dump session.findById("wnd[0]/usr"), 0
    Else
        log.WriteLine "  tbl=" & tbl.Id & " rows=" & tbl.RowCount & " visible=" & tbl.VisibleRowCount
        TableDump tbl.Id, 80
    End If
    On Error GoTo 0
Next

' ---------- 2) 자재변경 팝업 구조 (BSTKD 1630, 판매문서 없음) ----------
log.WriteLine "=============== ZSD600 조회 BSTKD=1630 → 품목 → CHANGE_MAT 팝업 (취소) ==============="
Tcode "/nZSD600"
With session
    .findById("wnd[0]/usr/radP_RBLIST").Select
    .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = "1630"
    .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
    .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.09.08"
    .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
    .findById("wnd[0]/tbar[1]/btn[8]").press
End With
WScript.Sleep 4000
Status
Dim gh, gi
Set gh = session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell")
Set gi = session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell")
log.WriteLine "-- header grid"
GridAll gh, 5
' 헤더 행 선택 → 품목 채워지는지
gh.setCurrentCell 0, "BSTKD"
gh.selectedRows = "0"
On Error Resume Next
gh.doubleClickCurrentCell
If Err.Number <> 0 Then log.WriteLine "  doubleClick err: " & Err.Description : Err.Clear
On Error GoTo 0
WScript.Sleep 1500
log.WriteLine "-- item grid after click"
Set gi = session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell")
GridAll gi, 10
If gi.RowCount = 0 Then
    gh.setCurrentCell 0, "BSTKD"
    gh.doubleClickCurrentCell
    WScript.Sleep 1500
    Status
    On Error Resume Next
    Set w1 = session.findById("wnd[1]")
    If Err.Number = 0 Then log.WriteLine "  dblclick popup=[" & w1.Text & "]" : Dump w1, 0 : w1.close : WScript.Sleep 500
    Err.Clear
    On Error GoTo 0
    Set gi = session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell")
    log.WriteLine "-- item grid after dblclick"
    GridAll gi, 10
End If
' 품목 1행 선택 후 CHANGE_MAT
If gi.RowCount > 0 Then
    gi.setCurrentCell 0, "MATNR"
    gi.selectedRows = "0"
    gi.pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 2000
    Status
    On Error Resume Next
    Set w1 = session.findById("wnd[1]")
    If Err.Number = 0 Then
        log.WriteLine "-- CHANGE_MAT popup"
        Dump w1, 0
        ' 취소: 닫기 (저장 안 함)
        w1.close
        WScript.Sleep 800
        Status
        Set w1 = session.findById("wnd[1]")
        If Err.Number = 0 Then log.WriteLine "  close 후 팝업 잔존=[" & w1.Text & "]" : Dump w1, 0 : w1.sendVKey 12 : WScript.Sleep 500
    Else
        log.WriteLine "  CHANGE_MAT 팝업 없음"
    End If
    Err.Clear
    On Error GoTo 0
End If
Status

Tcode "/n"
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "se11_chgmat done"
