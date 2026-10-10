' inspect_deep.vbs : 인클루드 버전이력 + SE11 테이블 필드 + Z33SD* 출고전기 검색 (읽기전용)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_deep.txt", True, True)
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
Sub ClosePopups()
    Dim n
    On Error Resume Next
    For n = 3 To 1 Step -1
        session.findById("wnd[" & n & "]").close
    Next
    Err.Clear
    On Error GoTo 0
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
Sub GridRows(grid, colsWanted, maxrows)
    Dim cols, c, r, line, maxr
    On Error Resume Next
    log.WriteLine "  rows=" & grid.RowCount
    Set cols = grid.ColumnOrder
    maxr = grid.RowCount - 1
    If maxr > maxrows - 1 Then maxr = maxrows - 1
    For r = 0 To maxr
        line = "  r" & r & ":"
        For c = 0 To cols.Count - 1
            If colsWanted = "" Or InStr("," & colsWanted & ",", "," & cols(c) & ",") > 0 Then
                line = line & " " & cols(c) & "=[" & grid.GetCellValue(r, cols(c)) & "]"
            End If
        Next
        log.WriteLine line
    Next
    On Error GoTo 0
End Sub
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
Sub DumpList()
    Dim page, usr, c, comp, cur, prev, nm, p1, p2, row, maxrow, d
    prev = ""
    For page = 1 To 40
        Set usr = session.findById("wnd[0]/usr")
        Set d = CreateObject("Scripting.Dictionary")
        maxrow = 0 : cur = ""
        For c = 0 To usr.Children.Count - 1
            Set comp = usr.Children.ElementAt(c)
            If comp.Type = "GuiLabel" Then
                nm = comp.Id : nm = Mid(nm, InStrRev(nm, "/") + 1)
                p1 = InStr(nm, "[") : p2 = InStr(nm, ",")
                If p1 = 0 Or p2 = 0 Then row = maxrow + 1 Else row = CInt(Mid(nm, p2 + 1, Len(nm) - p2 - 1))
                If row > maxrow Then maxrow = row
                If Not d.Exists(row) Then d.Add row, ""
                d(row) = d(row) & " " & comp.Text
            End If
        Next
        For row = 0 To maxrow
            If d.Exists(row) Then cur = cur & d(row) & vbCrLf
        Next
        If cur = prev Or cur = "" Then Exit For
        log.Write cur
        prev = cur
        session.findById("wnd[0]").sendVKey 82
        WScript.Sleep 600
    Next
End Sub

' ---------- 1) 인클루드 버전이력 ----------
Dim incs, inc, m, grid
incs = Array("Z33SDR600", "Z33SDR600_TOP", "Z33SDR600_F01", "Z33SDR600_MCLS", "Z33SDR600_MALV", "Z33SDR600_SEL", "Z33SDR600_LAY", "Z33SDR600_O01", "Z33SDR600_I01")
For Each inc In incs
    log.WriteLine "=============== 버전이력: " & inc & " ==============="
    Tcode "/nSE38"
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = inc
    Set m = FindMenu(session.findById("wnd[0]/mbar"), "버젼관리")
    If m Is Nothing Then
        log.WriteLine "  메뉴 없음"
    Else
        m.Select
        WScript.Sleep 2000
        ClosePopups
        On Error Resume Next
        Set grid = session.findById("wnd[0]/usr/cntlCONT/shellcont/shell/shellcont[0]/shell")
        If Err.Number <> 0 Then
            log.WriteLine "  grid 없음 screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " sbar=" & session.findById("wnd[0]/sbar").Text
            Err.Clear
        Else
            On Error GoTo 0
            GridRows grid, "VERSNO,REL,KORRNO,KORRTEXT,DATE,TIME,USR", 15
        End If
        On Error GoTo 0
    End If
Next

' ---------- 2) SE11 테이블 필드 ----------
Dim tabs, tb, tbl
tabs = Array("Z33SDT640", "Z33SDT641", "Z33SDT642", "Z33SDT643")
For Each tb In tabs
    log.WriteLine "=============== SE11 필드: " & tb & " ==============="
    Tcode "/nSE11"
    On Error Resume Next
    session.findById("wnd[0]/usr/radRSRD1-TBMA").Select
    session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").Text = tb
    session.findById("wnd[0]/tbar[1]/btn[7]").press
    WScript.Sleep 2000
    ClosePopups
    log.WriteLine "  screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " title=[" & session.findById("wnd[0]").Text & "] sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB02").Select
    WScript.Sleep 800
    Err.Clear
    Set tbl = FindType(session.findById("wnd[0]/usr"), "GuiTableControl")
    If tbl Is Nothing Then
        log.WriteLine "  테이블컨트롤 없음"
    Else
        log.WriteLine "  tbl=" & tbl.Id & " rows=" & tbl.RowCount & " visible=" & tbl.VisibleRowCount
        TableDump tbl.Id, 80
    End If
    On Error GoTo 0
Next

' ---------- 3) Z33SD* 전체 검색 ----------
Sub ScanAll(term, repLow, repHigh, label)
    log.WriteLine "=============== SCAN " & label & " [" & term & "] in " & repLow & "~" & repHigh & " ==============="
    Tcode "/nSE38"
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "RPR_ABAP_SOURCE_SCAN"
    session.findById("wnd[0]").sendVKey 8
    WScript.Sleep 1500
    With session
        .findById("wnd[0]/usr/txtSSTRING-LOW").Text = term
        .findById("wnd[0]/usr/chkP_REGEX").Selected = True
        .findById("wnd[0]/usr/ctxtREPNAME-LOW").Text = repLow
        .findById("wnd[0]/usr/ctxtREPNAME-HIGH").Text = repHigh
        .findById("wnd[0]/usr/txtPLUSMINU").Text = "0"
        .findById("wnd[0]/usr/chkINCLU").Selected = False
        .findById("wnd[0]/usr/chkCOMMENT").Selected = True
        .findById("wnd[0]/usr/radRB_CODE").Select
        .findById("wnd[0]").sendVKey 8
    End With
    WScript.Sleep 8000
    ClosePopups
    log.WriteLine "  screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    If session.Info.ScreenNumber <> 1000 Then DumpList
End Sub
ScanAll "GOODS_ISSUE|WS_DELIVERY_UPDATE|CONFIRM_DEC|BAPI_OUTB_DELIVERY|POST_GOODS|PGI", "Z33SD*", "", "출고전기/납품 BAPI"
ScanAll "Z33SDT64", "Z33SD*", "", "Z33SDT64x 사용 프로그램"
ScanAll "MATNR_O|OLD_MATNR|ORG_MATNR|MATNR_1C|CHG_MATNR|_HIST|ZCRMF02_RFC_MAT|E000\(Z33SDM\)|TEXT-E32|TEXT-M23", "Z33SDR600*", "", "자재변경/검증 관련"

Tcode "/n"
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "deep done"
