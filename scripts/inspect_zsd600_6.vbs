' inspect_zsd600_6.vbs : 주문상태 F4 값 목록 (읽기전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zsd600_6.txt", True, True)
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

Sub Dump(comp, depth)
    Dim c, k, line
    line = String(depth * 2, " ") & comp.Type & " | " & comp.Id
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

Sub F4Dump(fieldId, label)
    Dim grid, cols, c, r, line, w1
    log.WriteLine "=============== F4: " & label & " (" & fieldId & ") ==============="
    session.findById(fieldId).SetFocus
    session.findById("wnd[0]").sendVKey 4
    WScript.Sleep 1500
    On Error Resume Next
    Set w1 = session.findById("wnd[1]")
    If Err.Number <> 0 Then log.WriteLine "(팝업 없음) sbar=" & session.findById("wnd[0]/sbar").Text : Err.Clear : Exit Sub
    On Error GoTo 0
    log.WriteLine "popup title=[" & w1.Text & "]"
    Set grid = FindGrid(w1)
    If grid Is Nothing Then
        Dump w1, 0
    Else
        On Error Resume Next
        Set cols = grid.ColumnOrder
        line = "  cols:"
        For c = 0 To cols.Count - 1
            line = line & " " & cols(c) & "(" & grid.GetColumnTitles(cols(c))(0) & ")"
        Next
        log.WriteLine line
        For r = 0 To grid.RowCount - 1
            line = "  row" & r & ":"
            For c = 0 To cols.Count - 1
                line = line & " [" & grid.GetCellValue(r, cols(c)) & "]"
            Next
            log.WriteLine line
        Next
        On Error GoTo 0
    End If
    w1.close
    WScript.Sleep 500
End Sub

session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZSD600"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
F4Dump "wnd[0]/usr/ctxtS_STAT1-LOW", "주문상태(조회)"
session.findById("wnd[0]/usr/radP_RBSYNC").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800
F4Dump "wnd[0]/usr/ctxtS_STAT3-LOW", "주문상태(동기화)"
F4Dump "wnd[0]/usr/ctxtP_DBCON", "DB연결"

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect6 done"
