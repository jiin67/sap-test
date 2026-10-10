' KDD variant of S12_se16n_cols.vbs (2026-10-03): same script, system guard KDD. READ-ONLY.
' S12_se16n_cols.vbs  (S08 + column subset + output tag)
'   Query a table with SE16N (DISPLAY ONLY) and dump the result grid to TSV.
'   Never touches the "엔트리 유지보수" (edit) checkbox.
'   args: <table> <maxlines> <tag> <COL1,COL2,...|*> [<fieldDesc>=<low>[..<high>]] ...
'     field is matched against the SE16N selection list description (SCRTEXT_M),
'     e.g. "접수일자=2026.07.01..2026.07.15"
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, f, n0, i
Dim TBL, MAXLINES, TAG, COLS

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

Set connection = Nothing
For i = 0 To application.Children.Count - 1
    Set base = application.Children.ElementAt(i).Children.ElementAt(0)
    If base.Info.SystemName = "KDD" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then
    WScript.Echo "GUARD: no KDQ connection"
    WScript.Quit 99
End If
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDD" Then
    WScript.Echo "GUARD: " & base.Info.SystemName
    WScript.Quit 99
End If

Dim ARGS
Set ARGS = WScript.Arguments
If ARGS.Count < 4 Then
    WScript.Echo "usage: S08_se16n_query.vbs <table> <maxlines> [desc=low..high] ..."
    WScript.Quit 1
End If
TBL = UCase(ARGS(0))
MAXLINES = ARGS(1)
TAG = ARGS(2)
COLS = UCase(ARGS(3))

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\raw") Then fso.CreateFolder "C:\sap-test\results\raw"
Set log = fso.CreateTextFile("C:\sap-test\logs\S12k_" & TBL & "_" & TAG & ".txt", True, True)
Set f = fso.CreateTextFile("C:\sap-test\results\raw\S12k_" & TBL & "_" & TAG & ".tsv", True, True)
log.WriteLine "table=" & TBL & " maxlines=" & MAXLINES

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

' ---------- open SE16N and enter the table ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = TBL
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2500
If Err.Number <> 0 Then
    log.WriteLine "ERR open/table " & Err.Number & " " & Err.Description
    Err.Clear
End If
session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = MAXLINES
If Err.Number <> 0 Then
    log.WriteLine "ERR max_lines " & Err.Description
    Err.Clear
End If
' make sure edit mode stays OFF
Dim ed
Set ed = session.findById("wnd[0]/usr/chkGD-EDIT")
If Err.Number = 0 Then
    If ed.Selected Then
        log.WriteLine "WARN edit checkbox was ON - clearing"
        ed.Selected = False
    End If
End If
Err.Clear
On Error GoTo 0

' ---------- set one selection criterion by its description ----------
Function SetSelByDesc(desc, lowv, highv)
    Dim tc, r, txt, sb, pos, maxpos, pass, found
    found = False
    On Error Resume Next
    Set tc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC")
    If Err.Number <> 0 Then
        Err.Clear
        log.WriteLine "  ERR selection table control not found"
        SetSelByDesc = False
        Exit Function
    End If
    Set sb = tc.VerticalScrollbar
    maxpos = 0
    If Err.Number = 0 Then maxpos = sb.Maximum
    Err.Clear
    On Error GoTo 0

    pos = 0
    Do
        On Error Resume Next
        If True Then   ' always reset scroll (S08 bug: page 0 never revisited)
            session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VerticalScrollbar.Position = pos
            WScript.Sleep 400
            Set tc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC")
        End If
        Err.Clear
        On Error GoTo 0
        For r = 0 To 30
            On Error Resume Next
            txt = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & r & "]").Text
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            On Error GoTo 0
            If Trim(txt) = desc Then
                On Error Resume Next
                session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & r & "]").Text = lowv
                If Err.Number <> 0 Then
                    log.WriteLine "  ERR set LOW for " & desc & " : " & Err.Description
                    Err.Clear
                End If
                If highv <> "" Then
                    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3," & r & "]").Text = highv
                    If Err.Number <> 0 Then
                        log.WriteLine "  ERR set HIGH for " & desc & " : " & Err.Description
                        Err.Clear
                    End If
                End If
                On Error GoTo 0
                log.WriteLine "  set [" & desc & "] row " & r & " = " & lowv & " .. " & highv
                found = True
                Exit Do
            End If
        Next
        pos = pos + 15
    Loop While pos <= maxpos
    SetSelByDesc = found
End Function

' ---------- apply criteria from the command line ----------
Dim a, spec, kv, rng, okset
For a = 4 To ARGS.Count - 1
    spec = ARGS(a)
    kv = Split(spec, "=")
    If UBound(kv) >= 1 Then
        rng = Split(kv(1), "..")
        If UBound(rng) >= 1 Then
            okset = SetSelByDesc(Trim(kv(0)), Trim(rng(0)), Trim(rng(1)))
        Else
            okset = SetSelByDesc(Trim(kv(0)), Trim(kv(1)), "")
        End If
        If Not okset Then log.WriteLine "  WARN criterion not applied: " & spec
    End If
Next

' ---------- execute ----------
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 3000
If Err.Number <> 0 Then
    log.WriteLine "ERR execute " & Err.Number & " " & Err.Description
    Err.Clear
End If
log.WriteLine "after exec prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & _
              " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
Err.Clear
On Error GoTo 0

' ---------- find and dump the result grid ----------
Function FindGrid()
    Dim ids, j, g
    ids = Array("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell", _
                "wnd[0]/usr/cntlGRID1/shellcont/shell", _
                "wnd[0]/usr/cntlCONTAINER/shellcont/shell", _
                "wnd[0]/usr/cntlGRID/shellcont/shell", _
                "wnd[0]/usr/cntlCC_ALV/shellcont/shell")
    For j = 0 To UBound(ids)
        On Error Resume Next
        Set g = session.findById(ids(j))
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            log.WriteLine "grid at " & ids(j)
            Set FindGrid = g
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    Next
    Set FindGrid = Nothing
End Function

Sub TreeShells(comp, depth)
    Dim c, k, t
    On Error Resume Next
    t = comp.Type
    If t = "GuiShell" Then
        log.WriteLine String(depth, " ") & "SHELL " & Mid(comp.Id, InStr(comp.Id, "/wnd[")) & " sub=" & comp.SubType
        Err.Clear
    End If
    k = comp.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    For c = 0 To k - 1
        TreeShells comp.Children.ElementAt(c), depth + 1
    Next
End Sub

Dim g, ord, c, r, line, maxr, vis, winTop
Set g = FindGrid()
If g Is Nothing Then
    log.WriteLine "result grid NOT found - dumping shells"
    TreeShells session.findById("wnd[0]"), 0
    f.WriteLine "#GRID" & vbTab & "NOT_FOUND"
Else
    On Error Resume Next
    Set ord = g.ColumnOrder
    Dim cl, nc
    If COLS = "*" Then
        ReDim cl(ord.Count - 1)
        For c = 0 To ord.Count - 1
            cl(c) = ord(c)
        Next
    Else
        cl = Split(COLS, ",")
    End If
    nc = UBound(cl) + 1
    f.WriteLine "#TABLE" & vbTab & TBL
    f.WriteLine "#ROWCOUNT" & vbTab & g.RowCount
    f.WriteLine "#COLCOUNT" & vbTab & nc
    line = ""
    For c = 0 To nc - 1
        If c > 0 Then line = line & vbTab
        line = line & cl(c)
    Next
    f.WriteLine "#DATA"
    f.WriteLine line
    maxr = g.RowCount - 1
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
        For c = 0 To nc - 1
            If c > 0 Then line = line & vbTab
            line = line & Replace(Replace(g.GetCellValue(r, cl(c)), vbTab, " "), vbCrLf, " ")
            If Err.Number <> 0 Then Err.Clear
        Next
        f.WriteLine line
    Next
    Err.Clear
    On Error GoTo 0
    log.WriteLine "dumped rows=" & (maxr + 1)
    ' optional evidence screenshot: set env S12_SHOT=1 (grid scrolled to the first requested column)
    If CreateObject("WScript.Shell").ExpandEnvironmentStrings("%S12_SHOT%") = "1" Then
        On Error Resume Next
        g.FirstVisibleRow = 0
        If COLS <> "*" Then g.FirstVisibleColumn = cl(0)
        Err.Clear
        session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S12k_" & TBL & "_" & TAG & ".png", 2
        log.WriteLine "shot err=" & Err.Number
        If Err.Number <> 0 Then
            ' new session may open maximized (inherits the user's window) - HardCopy then fails; restore OUR window and retry
            Err.Clear
            session.findById("wnd[0]").restore
            WScript.Sleep 1000
            session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S12k_" & TBL & "_" & TAG & ".png", 2
            log.WriteLine "shot retry after restore err=" & Err.Number
        End If
        Err.Clear
        On Error GoTo 0
    End If
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
WScript.Echo "S12 " & TBL & " " & TAG & " done"
