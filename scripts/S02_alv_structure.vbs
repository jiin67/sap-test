' S02_alv_structure.vbs
'   TC-06 / TC-11 / TC-12 support: dump ALV column metadata and per-cell
'   type / editability for the first N rows, for both programs.
'   READ ONLY : runs the query, inspects the grid, never saves.
'   args: <dfrom> <dto> [rows]
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i
Dim OUTDIR, NROWS
OUTDIR = "C:\sap-test\results\raw\"
NROWS = 5

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
    WScript.Echo "GUARD: no KDQ connection found"
    WScript.Quit 99
End If
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then
    WScript.Echo "GUARD: system " & base.Info.SystemName
    WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\raw") Then fso.CreateFolder "C:\sap-test\results\raw"
Set log = fso.CreateTextFile("C:\sap-test\logs\S02_alv_structure.txt", True, True)

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

Dim D1, D2, ARGS
Set ARGS = WScript.Arguments
If ARGS.Count >= 2 Then
    D1 = ARGS(0)
    D2 = ARGS(1)
Else
    D1 = "2026.01.01"
    D2 = "2026.09.09"
End If
If ARGS.Count >= 3 Then NROWS = CInt(ARGS(2))
log.WriteLine "range " & D1 & " ~ " & D2 & " rows=" & NROWS

Function SetField(nm, val)
    Dim o
    On Error Resume Next
    Set o = session.findById("wnd[0]/usr/ctxt" & nm)
    If Err.Number <> 0 Then
        Err.Clear
        Set o = session.findById("wnd[0]/usr/txt" & nm)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        SetField = False
        Exit Function
    End If
    o.Text = val
    SetField = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' ---- new-program T-code is kept out of this file; read it from config_local.txt
Function NewTcode()
    Dim p, ts, ln, v
    v = ""
    p = "C:\sap-test\scripts\config_local.txt"
    On Error Resume Next
    If fso.FileExists(p) Then
        Set ts = fso.OpenTextFile(p, 1)
        Do While Not ts.AtEndOfStream
            ln = Trim(ts.ReadLine)
            If InStr(ln, "TCODE_NEW=") = 1 Then v = Trim(Mid(ln, 11))
        Loop
        ts.Close
    End If
    Err.Clear
    On Error GoTo 0
    NewTcode = v
End Function

Sub Probe(progname)
    Dim g, f, ord, ncol, nrow, c, r, line, path, v, maxr
    path = OUTDIR & "TC06_" & progname & "_struct.tsv"
    Set f = fso.CreateTextFile(path, True, True)
    f.WriteLine "#PROG" & vbTab & progname

    On Error Resume Next
    If progname = "ZCRMR052" Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 2500
    Else
        Dim tc
        tc = NewTcode()
        If tc <> "" Then
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tc
            session.findById("wnd[0]").sendVKey 0
            WScript.Sleep 2400
        Else
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
            session.findById("wnd[0]").sendVKey 0
            WScript.Sleep 1600
            session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "ZCRMR052S"
            session.findById("wnd[0]/usr/radRS38M-FUNC_EDIT").Select
            session.findById("wnd[0]/tbar[1]/btn[8]").press
            WScript.Sleep 3000
        End If
    End If
    session.findById("wnd[0]/usr/radPA_R01").Select
    session.findById("wnd[0]/usr/radPA_CHX").Select
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1300
    Err.Clear
    On Error GoTo 0
    Dim ok
    ok = SetField("SO_RCDT-LOW", D1)
    ok = SetField("SO_RCDT-HIGH", D2)

    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 9000
    Err.Clear

    f.WriteLine "#TITLE" & vbTab & session.findById("wnd[0]").Text
    f.WriteLine "#HEADER_REMARK1" & vbTab & session.findById("wnd[0]/usr/txtGS_HEADER-REMARK1").Text
    f.WriteLine "#HEADER_REMARK2" & vbTab & session.findById("wnd[0]/usr/txtGS_HEADER-REMARK2").Text
    Err.Clear

    ' toolbar buttons of screen 100
    Dim b, j
    For j = 0 To 40
        Set b = session.findById("wnd[0]/tbar[1]/btn[" & j & "]")
        If Err.Number = 0 Then
            f.WriteLine "#TBAR" & vbTab & j & vbTab & b.Text & vbTab & b.Tooltip
        End If
        Err.Clear
    Next

    Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
    If Err.Number <> 0 Then
        Err.Clear
        f.WriteLine "#GRID" & vbTab & "NOT_FOUND"
        f.Close
        Exit Sub
    End If
    nrow = g.RowCount
    Set ord = g.ColumnOrder
    ncol = ord.Count
    f.WriteLine "#ROWCOUNT" & vbTab & nrow
    f.WriteLine "#COLCOUNT" & vbTab & ncol

    ' grid-level toolbar (ALV buttons)
    Dim tb, tcnt
    tcnt = g.GetToolbarButtonCount
    If Err.Number = 0 Then
        f.WriteLine "#ALVTBCOUNT" & vbTab & tcnt
        For j = 0 To tcnt - 1
            f.WriteLine "#ALVTB" & vbTab & j & vbTab & g.GetToolbarButtonId(j) & vbTab & _
                        g.GetToolbarButtonType(j) & vbTab & g.GetToolbarButtonText(j)
            Err.Clear
        Next
    End If
    Err.Clear

    ' column metadata
    Dim titles, tt, k
    For c = 0 To ncol - 1
        line = "#COL" & vbTab & c & vbTab & ord(c)
        tt = ""
        Set titles = g.GetColumnTitles(ord(c))
        If Err.Number = 0 Then
            For k = 0 To titles.Count - 1
                If tt <> "" Then tt = tt & "|"
                tt = tt & titles(k)
            Next
        End If
        Err.Clear
        line = line & vbTab & "titles=" & tt
        v = g.GetColumnDataType(ord(c))
        If Err.Number = 0 Then line = line & vbTab & "dtype=" & v
        Err.Clear
        v = g.GetColumnTooltip(ord(c))
        If Err.Number = 0 Then line = line & vbTab & "tip=" & v
        Err.Clear
        f.WriteLine line
    Next

    ' per-cell metadata for first NROWS rows
    maxr = nrow - 1
    If maxr > NROWS - 1 Then maxr = NROWS - 1
    For r = 0 To maxr
        For c = 0 To ncol - 1
            line = "#CELL" & vbTab & r & vbTab & ord(c)
            v = g.GetCellValue(r, ord(c))
            If Err.Number = 0 Then line = line & vbTab & "val=[" & v & "]"
            Err.Clear
            v = g.GetCellType(r, ord(c))
            If Err.Number = 0 Then line = line & vbTab & "type=" & v
            Err.Clear
            v = g.GetCellChangeable(r, ord(c))
            If Err.Number = 0 Then line = line & vbTab & "chg=" & v
            Err.Clear
            v = g.GetCellState(r, ord(c))
            If Err.Number = 0 Then line = line & vbTab & "state=" & v
            Err.Clear
            f.WriteLine line
        Next
    Next
    On Error GoTo 0
    f.Close
    log.WriteLine progname & " rows=" & nrow & " cols=" & ncol & " -> " & path

    ' back to menu
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1300
    Dim w
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    On Error GoTo 0
End Sub

Probe "ZCRMR052"
Probe "ZCRMR052S"

log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "S02_alv_structure done"
