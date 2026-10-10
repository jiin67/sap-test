' S05_change_layout.vbs
'   TC-06 decisive check: open the ALV "레이아웃 변경" dialog (&OL0) on each program
'   and dump BOTH lists - displayed columns and the hidden column pool.
'   READ ONLY : cancels the dialog with F12, never saves a layout.
'   args: <dfrom> <dto>
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, f, n0, i

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

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S05_change_layout.txt", True, True)
Set f = fso.CreateTextFile("C:\sap-test\results\raw\TC06_change_layout.tsv", True, True)

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
    D1 = "2026.07.01"
    D2 = "2026.07.01"
End If

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

Function RunQuery(progname)
    Dim g, ok
    On Error Resume Next
    If progname = "ZCRMR052" Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 2200
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
            WScript.Sleep 2500
        End If
    End If
    session.findById("wnd[0]/usr/radPA_R01").Select
    session.findById("wnd[0]/usr/radPA_CHX").Select
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1100
    Err.Clear
    On Error GoTo 0
    ok = SetField("SO_RCDT-LOW", D1)
    ok = SetField("SO_RCDT-HIGH", D2)
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 1500
    Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
    If Err.Number <> 0 Then
        Err.Clear
        Set RunQuery = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set RunQuery = g
End Function

' dump every GridView inside wnd[1] with all of its rows

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

Sub DumpDialogGrids(progname, tag)
    Dim w
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        f.WriteLine progname & vbTab & tag & vbTab & "NO_POPUP"
        Exit Sub
    End If
    f.WriteLine progname & vbTab & tag & vbTab & "TITLE" & vbTab & w.Text
    On Error GoTo 0
    WalkGrids w, progname, tag
End Sub

Sub WalkGrids(root, progname, tag)
    Dim c, k, comp, t, sb, g2, ord, r, cc, line
    On Error Resume Next
    k = root.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For c = 0 To k - 1
        Set comp = root.Children.ElementAt(c)
        t = comp.Type
        If t = "GuiShell" Then
            sb = comp.SubType
            If sb = "GridView" Then
                Set g2 = comp
                Set ord = g2.ColumnOrder
                f.WriteLine progname & vbTab & tag & vbTab & "GRID" & vbTab & _
                            Mid(comp.Id, InStr(comp.Id, "/wnd[1]")) & vbTab & _
                            "rows=" & g2.RowCount & " cols=" & ord.Count
                For r = 0 To g2.RowCount - 1
                    line = ""
                    For cc = 0 To ord.Count - 1
                        If cc > 0 Then line = line & " | "
                        line = line & ord(cc) & "=" & g2.GetCellValue(r, ord(cc))
                    Next
                    f.WriteLine progname & vbTab & tag & vbTab & "ROW" & vbTab & r & vbTab & line
                    Err.Clear
                    If r > 300 Then Exit For
                Next
            Else
                f.WriteLine progname & vbTab & tag & vbTab & "SHELL" & vbTab & _
                            Mid(comp.Id, InStr(comp.Id, "/wnd[1]")) & vbTab & sb
            End If
            Err.Clear
        ElseIf t = "GuiTab" Then
            f.WriteLine progname & vbTab & tag & vbTab & "TAB" & vbTab & _
                        Mid(comp.Id, InStr(comp.Id, "/wnd[1]")) & vbTab & comp.Text
            Err.Clear
        End If
        WalkGrids comp, progname, tag
    Next
    Err.Clear
End Sub

Sub ClosePopups()
    Dim w, g
    g = 0
    Do
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        On Error GoTo 0
        w.sendVKey 12
        WScript.Sleep 800
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number = 0 Then
            w.close
            WScript.Sleep 700
        End If
        Err.Clear
        On Error GoTo 0
        g = g + 1
    Loop While g < 4
End Sub

Sub Reset()
    On Error Resume Next
    ClosePopups
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1100
    ClosePopups
    On Error GoTo 0
End Sub

Sub Probe(progname)
    Dim g, ord, c, fc
    f.WriteLine "############" & vbTab & progname
    Set g = RunQuery(progname)
    If g Is Nothing Then
        f.WriteLine progname & vbTab & "NO_GRID"
        Reset
        Exit Sub
    End If
    On Error Resume Next
    Set ord = g.ColumnOrder
    f.WriteLine progname & vbTab & "#VISIBLECOLS" & vbTab & ord.Count
    Err.Clear
    On Error GoTo 0

    ' try the standard ALV change-layout function codes in order
    Dim codes, ci, opened
    codes = Array("&OL0", "&OAD", "&LAYOUT")
    opened = False
    For ci = 0 To UBound(codes)
        On Error Resume Next
        g.pressToolbarButton codes(ci)
        WScript.Sleep 1800
        If Err.Number <> 0 Then
            log.WriteLine progname & " " & codes(ci) & " err " & Err.Number & " " & Err.Description
            Err.Clear
        End If
        Dim w
        Set w = session.findById("wnd[1]")
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            log.WriteLine progname & " opened with " & codes(ci)
            f.WriteLine progname & vbTab & "#OPENEDBY" & vbTab & codes(ci)
            DumpDialogGrids progname, "CHGLAYOUT"
            ClosePopups
            opened = True
            Exit For
        End If
        Err.Clear
        On Error GoTo 0
    Next
    If Not opened Then
        f.WriteLine progname & vbTab & "#CHGLAYOUT" & vbTab & "COULD_NOT_OPEN"
        log.WriteLine progname & " could not open change-layout dialog"
    End If
    Reset
End Sub

Probe "ZCRMR052"
Probe "ZCRMR052S"

f.Close
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "S05_change_layout done"
