' S01_compare_query.vbs
'   Runs the same selection conditions on ZCRM052 (ZCRMR052) and on ZCRMR052S,
'   dumps each result ALV to a TSV file for external diffing (compare_tsv.py).
'   READ ONLY : executes queries only, never saves.
'   args: <dfrom> <dto> [caseFilter] [rowCap]
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i
Dim OUTDIR, ROWCAP
OUTDIR = "C:\sap-test\results\raw\"
ROWCAP = 500

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

' ---- pick the KDQ connection explicitly (a KDD connection may also be open)
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
Set log = fso.CreateTextFile("C:\sap-test\logs\S01_compare_query.txt", True, True)
log.WriteLine "system=" & base.Info.SystemName & " client=" & base.Info.Client & " user=" & base.Info.User

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
    If connection.Children.Count <= n0 Then
        log.WriteLine "FAIL create session"
        log.Close
        WScript.Quit 1
    End If
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    WScript.Sleep 1500
End If
log.WriteLine "work session=" & session.Id

'---------------------------------------------------------------- helpers
Function FieldExists(nm)
    Dim o
    On Error Resume Next
    Set o = session.findById("wnd[0]/usr/" & nm)
    FieldExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

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

Sub ClearDates()
    Dim r
    r = SetField("SO_RCDT-LOW", "")
    r = SetField("SO_RCDT-HIGH", "")
    r = SetField("SO_PRDRE-LOW", "")
    r = SetField("SO_PRDRE-HIGH", "")
End Sub

Sub OpenOld()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2200
End Sub

Sub OpenNew()
    Dim tc
    tc = NewTcode()
    If tc <> "" Then
        ' run through the dedicated T-code : correct GUI status / toolbar
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tc
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 2400
    Else
        ' fallback : execute the program from SE38 (adds ABAP editor buttons to tbar[1])
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 1600
        session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "ZCRMR052S"
        session.findById("wnd[0]/usr/radRS38M-FUNC_EDIT").Select
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        WScript.Sleep 2500
    End If
End Sub

Sub ApplySpec(spec)
    Dim parts, p, kv
    If spec = "" Then Exit Sub
    parts = Split(spec, ";")
    For Each p In parts
        If Trim(p) <> "" Then
            kv = Split(p, "=")
            If UBound(kv) >= 1 Then
                If Not SetField(Trim(kv(0)), Trim(kv(1))) Then
                    log.WriteLine "    WARN cannot set " & kv(0)
                End If
            End If
        End If
    Next
End Sub

Function TitleCount(t)
    Dim a, b
    a = InStrRev(t, "(")
    b = InStrRev(t, ")")
    If a > 0 And b > a Then
        TitleCount = Replace(Trim(Mid(t, a + 1, b - a - 1)), ",", "")
    Else
        TitleCount = ""
    End If
End Function

Function SelState()
    Dim s
    s = ""
    If FieldExists("ctxtSO_VKBUR-LOW") Then s = s & "VKBUR=Y;" Else s = s & "VKBUR=N;"
    If FieldExists("ctxtSO_KUNNR-LOW") Then s = s & "KUNNR=Y;" Else s = s & "KUNNR=N;"
    If FieldExists("ctxtSO_RCDT-LOW") Then s = s & "RCDT=Y;" Else s = s & "RCDT=N;"
    If FieldExists("ctxtSO_PRDRE-LOW") Then s = s & "PRDRE=Y;" Else s = s & "PRDRE=N;"
    SelState = s
End Function

Function DumpGrid(path, caseid, progname, spec, selstate)
    Dim g, f, c, r, line, v, nrow, ncol, ord, titles, tt, k, maxr, vis, winTop
    Set f = fso.CreateTextFile(path, True, True)
    f.WriteLine "#CASE" & vbTab & caseid
    f.WriteLine "#PROG" & vbTab & progname
    f.WriteLine "#SPEC" & vbTab & spec
    f.WriteLine "#SELSTATE" & vbTab & selstate
    On Error Resume Next
    f.WriteLine "#SCREEN" & vbTab & session.Info.ScreenNumber
    f.WriteLine "#TITLE" & vbTab & session.findById("wnd[0]").Text
    f.WriteLine "#TITLECOUNT" & vbTab & TitleCount(session.findById("wnd[0]").Text)
    f.WriteLine "#SBAR" & vbTab & session.findById("wnd[0]/sbar").Text
    Err.Clear

    Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
    If Err.Number <> 0 Then
        Err.Clear
        f.WriteLine "#GRID" & vbTab & "NOT_FOUND"
        f.Close
        DumpGrid = -1
        Exit Function
    End If
    nrow = g.RowCount
    Set ord = g.ColumnOrder
    ncol = ord.Count
    f.WriteLine "#ROWCOUNT" & vbTab & nrow
    f.WriteLine "#COLCOUNT" & vbTab & ncol

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
        f.WriteLine line & vbTab & tt
    Next

    line = ""
    For c = 0 To ncol - 1
        If c > 0 Then line = line & vbTab
        line = line & ord(c)
    Next

    maxr = nrow - 1
    If maxr > ROWCAP - 1 Then maxr = ROWCAP - 1
    f.WriteLine "#DUMPEDROWS" & vbTab & (maxr + 1)
    f.WriteLine "#DATA"
    f.WriteLine line

    vis = g.VisibleRowCount
    If Err.Number <> 0 Or vis < 1 Then
        vis = 1
        Err.Clear
    End If
    winTop = -1

    For r = 0 To maxr
        ' the ALV only holds the visible window client-side; scroll when needed
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
        For c = 0 To ncol - 1
            v = g.GetCellValue(r, ord(c))
            If Err.Number <> 0 Then
                v = "<ERR>"
                Err.Clear
            End If
            If c > 0 Then line = line & vbTab
            line = line & Replace(Replace(v, vbTab, " "), vbCrLf, " ")
        Next
        f.WriteLine line
    Next
    On Error GoTo 0
    f.Close
    DumpGrid = nrow
End Function

Sub Reset()
    Dim w
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1100
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    On Error GoTo 0
End Sub

Sub RunOne(caseid, isNew, org, recv, spec)
    Dim progname, path, ss, rc, t0
    If isNew Then
        progname = "ZCRMR052S"
    Else
        progname = "ZCRMR052"
    End If
    path = OUTDIR & caseid & "_" & progname & ".tsv"
    log.WriteLine "  -> " & caseid & " " & progname
    t0 = Timer
    On Error Resume Next
    If isNew Then
        OpenNew
    Else
        OpenOld
    End If
    If Err.Number <> 0 Then
        log.WriteLine "    ERR open " & Err.Description
        Err.Clear
    End If
    session.findById("wnd[0]/usr/" & org).Select
    session.findById("wnd[0]/usr/" & recv).Select
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1100
    If Err.Number <> 0 Then
        log.WriteLine "    ERR radio " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    ClearDates
    ApplySpec spec
    ss = SelState()
    log.WriteLine "    selstate " & ss
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 1500
    If Err.Number <> 0 Then
        log.WriteLine "    ERR exec " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    rc = DumpGrid(path, caseid, progname, org & "+" & recv & ";" & spec, ss)
    log.WriteLine "    rows=" & rc & " sec=" & FormatNumber(Timer - t0, 1)
    Reset
End Sub

Sub RunCase(caseid, org, recv, spec)
    log.WriteLine "=========== " & caseid & " " & org & "+" & recv & " " & spec
    RunOne caseid, False, org, recv, spec
    RunOne caseid, True, org, recv, spec
End Sub

'---------------------------------------------------------------- cases
Dim D1, D2, ARGS, ONLY
Set ARGS = WScript.Arguments
If ARGS.Count >= 2 Then
    D1 = ARGS(0)
    D2 = ARGS(1)
Else
    D1 = "2026.07.01"
    D2 = "2026.07.15"
End If
If ARGS.Count >= 3 Then
    ONLY = UCase(ARGS(2))
Else
    ONLY = "ALL"
End If
If ARGS.Count >= 4 Then ROWCAP = CInt(ARGS(3))
log.WriteLine "date range " & D1 & " ~ " & D2 & "  filter=" & ONLY & "  rowcap=" & ROWCAP

If ONLY = "ALL" Or ONLY = "TC01" Then RunCase "TC01", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2
If ONLY = "ALL" Or ONLY = "TC02" Then RunCase "TC02", "radPA_R02", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2
If ONLY = "ALL" Or ONLY = "TC03" Then RunCase "TC03", "radPA_R01", "radPA_NCH", "SO_PRDRE-LOW=" & D1 & ";SO_PRDRE-HIGH=" & D2
If ONLY = "ALL" Or ONLY = "TC04" Then RunCase "TC04Y", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2 & ";SO_PRTY-LOW=Y"
If ONLY = "ALL" Or ONLY = "TC04" Then RunCase "TC04N", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2 & ";SO_PRTY-LOW=N"
If ONLY = "ALL" Or ONLY = "TC05" Then RunCase "TC05O", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2 & ";SO_MULTK-LOW=O"
If ONLY = "ALL" Or ONLY = "TC05" Then RunCase "TC05X", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2 & ";SO_MULTK-LOW=X"

log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "S01_compare_query done"
