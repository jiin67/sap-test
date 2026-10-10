' S07_find_testdata.vbs
'   Find comparison / save test data with SE16N, following the data search guide.
'   READ ONLY : SE16N display only. Never uses SE16N change mode (&SAP_EDIT).
'   args: <table> [maxhits]
'   The script first dumps the SE16N screen structure, then the result grid.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, f, n0, i
Dim TBL, MAXHITS

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

Set connection = Nothing
For i = 0 To application.Children.Count - 1
    Set base = application.Children.ElementAt(i).Children.ElementAt(0)
    If base.Info.SystemName = "KDQ" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then
    WScript.Echo "GUARD: no KDQ connection"
    WScript.Quit 99
End If
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then
    WScript.Echo "GUARD: " & base.Info.SystemName
    WScript.Quit 99
End If

Dim ARGS
Set ARGS = WScript.Arguments
If ARGS.Count >= 1 Then
    TBL = UCase(ARGS(0))
Else
    TBL = "ZCRM_RECPT"
End If
If ARGS.Count >= 2 Then
    MAXHITS = ARGS(1)
Else
    MAXHITS = "50"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\raw") Then fso.CreateFolder "C:\sap-test\results\raw"
Set log = fso.CreateTextFile("C:\sap-test\logs\S07_find_testdata_" & TBL & ".txt", True, True)
Set f = fso.CreateTextFile("C:\sap-test\results\raw\SE16N_" & TBL & ".tsv", True, True)
log.WriteLine "table=" & TBL & " maxhits=" & MAXHITS

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
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    WScript.Sleep 1500
End If
log.WriteLine "work session=" & session.Id

Sub Status(tag)
    On Error Resume Next
    log.WriteLine "[" & tag & "] prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & _
                  " tcode=" & session.Info.Transaction & _
                  " title=[" & session.findById("wnd[0]").Text & "]" & _
                  " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    Err.Clear
    On Error GoTo 0
End Sub

Sub Tree(comp, depth)
    Dim c, k, line, st, t
    On Error Resume Next
    t = comp.Type
    line = String(depth * 2, " ") & t & " | " & Mid(comp.Id, InStr(comp.Id, "/wnd["))
    If t = "GuiShell" Then
        st = comp.SubType
        If Err.Number = 0 Then line = line & " | sub=" & st
        Err.Clear
    End If
    If t = "GuiTextField" Or t = "GuiCTextField" Or t = "GuiLabel" _
       Or t = "GuiCheckBox" Or t = "GuiRadioButton" Or t = "GuiButton" Then
        line = line & " | text=[" & comp.Text & "]"
    End If
    Err.Clear
    log.WriteLine line
    k = comp.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    For c = 0 To k - 1
        Tree comp.Children.ElementAt(c), depth + 1
    Next
End Sub

Function FindGrid()
    Dim ids, j, g
    ids = Array("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell", _
                "wnd[0]/usr/cntlGRID1/shellcont/shell", _
                "wnd[0]/usr/cntlCONTAINER/shellcont/shell", _
                "wnd[0]/usr/cntlGRID/shellcont/shell")
    For j = 0 To UBound(ids)
        On Error Resume Next
        Set g = session.findById(ids(j))
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            log.WriteLine "grid found at " & ids(j)
            Set FindGrid = g
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    Next
    Set FindGrid = Nothing
End Function

Sub DumpGrid(g, tag)
    Dim ord, c, r, line, maxr
    On Error Resume Next
    Set ord = g.ColumnOrder
    f.WriteLine "#TAG" & vbTab & tag
    f.WriteLine "#ROWCOUNT" & vbTab & g.RowCount
    f.WriteLine "#COLCOUNT" & vbTab & ord.Count
    line = ""
    For c = 0 To ord.Count - 1
        If c > 0 Then line = line & vbTab
        line = line & ord(c)
    Next
    f.WriteLine "#DATA"
    f.WriteLine line
    maxr = g.RowCount - 1
    If maxr > 199 Then maxr = 199
    Dim vis, winTop
    vis = g.VisibleRowCount
    If Err.Number <> 0 Or vis < 1 Then
        vis = 1
        Err.Clear
    End If
    winTop = -1
    For r = 0 To maxr
        If r < winTop Or r >= winTop + vis Then
            g.FirstVisibleRow = r
            If Err.Number <> 0 Then Err.Clear
            winTop = g.FirstVisibleRow
            If Err.Number <> 0 Then
                winTop = r
                Err.Clear
            End If
        End If
        line = ""
        For c = 0 To ord.Count - 1
            If c > 0 Then line = line & vbTab
            line = line & Replace(g.GetCellValue(r, ord(c)), vbTab, " ")
            If Err.Number <> 0 Then Err.Clear
        Next
        f.WriteLine line
    Next
    Err.Clear
    On Error GoTo 0
End Sub

' ---- open SE16N ----
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2200
If Err.Number <> 0 Then
    log.WriteLine "ERR open SE16N " & Err.Number & " " & Err.Description
    Err.Clear
End If
On Error GoTo 0
Status "SE16N initial"
log.WriteLine "--- usr tree (initial) ---"
Tree session.findById("wnd[0]/usr"), 0

' ---- table + max hits ----
On Error Resume Next
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = TBL
If Err.Number <> 0 Then
    log.WriteLine "ERR GD-TAB " & Err.Number & " " & Err.Description
    Err.Clear
End If
session.findById("wnd[0]/usr/txtGD-MAXHITS").Text = MAXHITS
If Err.Number <> 0 Then
    Err.Clear
    session.findById("wnd[0]/usr/ctxtGD-MAXHITS").Text = MAXHITS
    Err.Clear
End If
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2500
Err.Clear
On Error GoTo 0
Status "after table"
log.WriteLine "--- wnd[0] tree (after table) ---"
Tree session.findById("wnd[0]"), 0

' ---- the selection-criteria grid (field list) ----
Dim sel
Set sel = FindGrid()
If Not sel Is Nothing Then
    DumpGrid sel, "SELECTION_FIELDS"
Else
    log.WriteLine "selection grid not found"
End If

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
Dim w
Set w = session.findById("wnd[1]")
If Err.Number = 0 Then w.sendVKey 0
Err.Clear
On Error GoTo 0
f.Close
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "S07_find_testdata done"
