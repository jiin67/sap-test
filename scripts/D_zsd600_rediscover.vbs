' D_zsd600_rediscover.vbs : re-baseline modified Z33SDR600 (READ ONLY, no DB change)
'  1) include version history (confirm change landed)
'  2) SE11 Z33SDT644 fields (material-change history table)
'  3) ZSD600 open BSTKD 2190 -> header/item dump, result-screen toolbar buttons of both grids
'  4) double-click header -> item ATP coloring (HardCopy evidence)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder("C:\sap-test\results\shots")
Set log = fso.CreateTextFile("C:\sap-test\logs\D_zsd600_rediscover.txt", True, True)
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
log.WriteLine "System=" & session.Info.SystemName & " Client=" & session.Info.Client & " User=" & session.Info.User & " RunAt=" & Now

Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Function HasPopup()
    On Error Resume Next
    HasPopup = False
    Dim w : Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then HasPopup = True
    Err.Clear
    On Error GoTo 0
End Function
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
Sub GridToolbar(gid, label)
    Dim g, b
    On Error Resume Next
    Set g = session.findById(gid)
    If Err.Number <> 0 Then log.WriteLine "  (" & label & " grid missing)" : Err.Clear : Exit Sub
    log.WriteLine "  " & label & " grid rows=" & g.RowCount & " cols=" & g.ColumnCount & " toolbarBtns=" & g.ToolbarButtonCount
    For b = 0 To g.ToolbarButtonCount - 1
        log.WriteLine "    tb" & b & " id=[" & g.GetToolbarButtonId(b) & "] type=" & g.GetToolbarButtonType(b)
    Next
    On Error GoTo 0
End Sub
Sub TableDump(tblId, maxrows)
    Dim pos, r, c, line, cell, visible, lastFirst, first, tbl
    lastFirst = "###"
    On Error Resume Next
    Set tbl = session.findById(tblId)
    If Err.Number <> 0 Then log.WriteLine "  (tbl missing " & tblId & ")" : Err.Clear : Exit Sub
    visible = tbl.VisibleRowCount
    For pos = 0 To maxrows Step visible
        Set tbl = session.findById(tblId)
        tbl.verticalScrollbar.position = pos
        WScript.Sleep 300
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
            If line <> "" Then log.WriteLine "   " & line
        Next
        If first = lastFirst Or first = "" Then Exit For
        lastFirst = first
    Next
    On Error GoTo 0
End Sub

' ---------- 1) include version history ----------
Dim incs, inc, m, grid, cols, cc, rr, line
incs = Array("Z33SDR600", "Z33SDR600_TOP", "Z33SDR600_F01", "Z33SDR600_MCLS", "Z33SDR600_MALV", "Z33SDR600_SEL", "Z33SDR600_O01", "Z33SDR600_I01")
For Each inc In incs
    Tcode "/nSE38"
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = inc
    Set m = FindMenu(session.findById("wnd[0]/mbar"), ChrW(&HBC84) & ChrW(&HC838) & ChrW(&HAD00) & ChrW(&HB9AC))  ' beojeon-gwanri
    If m Is Nothing Then
        log.WriteLine inc & " : version-menu not found"
    Else
        m.Select
        WScript.Sleep 1800
        ClosePopups
        On Error Resume Next
        Set grid = session.findById("wnd[0]/usr/cntlCONT/shellcont/shell/shellcont[0]/shell")
        If Err.Number <> 0 Then
            log.WriteLine inc & " : no version grid (screen " & session.Info.ScreenNumber & ")"
            Err.Clear
        Else
            Set cols = grid.ColumnOrder
            line = inc & " :"
            For rr = 0 To grid.RowCount - 1
                Dim vv : vv = "" : vv = grid.GetCellValue(rr, "VERSNO")
                line = line & " [v" & vv & " KORR=" & grid.GetCellValue(rr, "KORRNO") & " " & grid.GetCellValue(rr, "DATE") & " " & grid.GetCellValue(rr, "USR") & "]"
            Next
            log.WriteLine line
        End If
        On Error GoTo 0
    End If
Next

' ---------- 2) SE11 Z33SDT644 + Z33SDT643 fields ----------
Dim tabs, tb, tbl
tabs = Array("Z33SDT644", "Z33SDT643")
For Each tb In tabs
    Tcode "/nSE11"
    On Error Resume Next
    session.findById("wnd[0]/usr/radRSRD1-TBMA").Select
    session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").Text = tb
    session.findById("wnd[0]").sendVKey 7
    WScript.Sleep 2000
    ClosePopups
    log.WriteLine "=== SE11 " & tb & " screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").Select
    WScript.Sleep 500
    Err.Clear
    Set tbl = FindType(session.findById("wnd[0]/usr"), "GuiTableControl")
    If Not tbl Is Nothing Then TableDump tbl.Id, 60
    On Error GoTo 0
Next

' ---------- 3) ZSD600 open BSTKD 2190 ----------
Tcode "/nZSD600"
With session
    .findById("wnd[0]/usr/radP_RBLIST").Select
    .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = "2190"
    .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
    .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
    .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
    .findById("wnd[0]/tbar[1]/btn[8]").press
End With
WScript.Sleep 3500
log.WriteLine "=== ZSD600 BSTKD=2190 screen=" & session.Info.ScreenNumber & " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
If session.Info.ScreenNumber = 100 Then
    Dim gh : Set gh = session.findById(HG)
    log.WriteLine "  header rows=" & gh.RowCount
    Dim r2
    For r2 = 0 To gh.RowCount - 1
        If r2 > 4 Then Exit For
        log.WriteLine "  H" & r2 & ": ID=" & gh.GetCellValue(r2,"ID") & " BSTKD=" & gh.GetCellValue(r2,"BSTKD") & " VERSION=" & gh.GetCellValue(r2,"VERSION") & " STATU=" & gh.GetCellValue(r2,"STATU") & " VBELN=" & gh.GetCellValue(r2,"VBELN") & " VBELN_VL=" & gh.GetCellValue(r2,"VBELN_VL") & " VDATU=" & gh.GetCellValue(r2,"VDATU") & " KUNNR=" & gh.GetCellValue(r2,"KUNNR")
    Next
    GridToolbar HG, "HEADER"
    ' double-click header 0 -> item ATP display
    gh.setCurrentCell 0, "BSTKD"
    gh.selectedRows = "0"
    On Error Resume Next
    gh.doubleClickCurrentCell
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 2500
    ClosePopups
    Dim gi : Set gi = session.findById(IG)
    log.WriteLine "  item rows=" & gi.RowCount
    For r2 = 0 To gi.RowCount - 1
        If r2 > 6 Then Exit For
        log.WriteLine "  I" & r2 & ": SEQNO=" & gi.GetCellValue(r2,"SEQNO") & " MATNR=" & gi.GetCellValue(r2,"MATNR") & " MAKTX=" & gi.GetCellValue(r2,"MAKTX") & " ZMENG=" & gi.GetCellValue(r2,"ZMENG") & " MEINS=" & gi.GetCellValue(r2,"MEINS") & " LGORT=" & gi.GetCellValue(r2,"LGORT")
    Next
    GridToolbar IG, "ITEM"
    ' HardCopy evidence of item grid coloring
    On Error Resume Next
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\D_2190_items.png", 2
    Err.Clear
    On Error GoTo 0
    log.WriteLine "  (screenshot: results\shots\D_2190_items.png)"
End If

Tcode "/n"
log.WriteLine "done"
log.Close
If Not HasPopup() Then session.findById("wnd[0]").close
WScript.Echo "rediscover done"
