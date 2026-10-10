' S04_layout_catalog.vbs
'   TC-06: open the ALV "레이아웃 변경" dialog on each program and dump the FULL
'   column catalog, including columns that are currently hidden.
'   READ ONLY : opens the dialog and closes it with Cancel, never saves a layout.
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
If Not fso.FolderExists("C:\sap-test\results\raw") Then fso.CreateFolder "C:\sap-test\results\raw"
Set log = fso.CreateTextFile("C:\sap-test\logs\S04_layout_catalog.txt", True, True)
Set f = fso.CreateTextFile("C:\sap-test\results\raw\TC06_layout_catalog.tsv", True, True)

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
    D2 = "2026.07.02"
End If
log.WriteLine "range " & D1 & " ~ " & D2

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

' dump the ALV grid toolbar so we know which buttons exist

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

Sub DumpAlvToolbar(g, progname)
    Dim j, cnt
    On Error Resume Next
    cnt = g.GetToolbarButtonCount
    If Err.Number <> 0 Then
        Err.Clear
        f.WriteLine progname & vbTab & "#ALVTB" & vbTab & "NOT_AVAILABLE"
        Exit Sub
    End If
    For j = 0 To cnt - 1
        f.WriteLine progname & vbTab & "#ALVTB" & vbTab & j & vbTab & _
                    g.GetToolbarButtonId(j) & vbTab & g.GetToolbarButtonType(j) & vbTab & _
                    g.GetToolbarButtonText(j) & vbTab & g.GetToolbarButtonTooltip(j)
        Err.Clear
    Next
    On Error GoTo 0
End Sub

' walk the menubar and log every entry so the layout menu path can be found
Sub DumpMenus(progname)
    Dim m0, m1, m2, a, b, c
    On Error Resume Next
    Set m0 = session.findById("wnd[0]/mbar")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For a = 0 To m0.Children.Count - 1
        Set m1 = m0.Children.ElementAt(a)
        f.WriteLine progname & vbTab & "#MENU1" & vbTab & a & vbTab & m1.Text
        For b = 0 To m1.Children.Count - 1
            Set m2 = m1.Children.ElementAt(b)
            f.WriteLine progname & vbTab & "#MENU2" & vbTab & a & "/" & b & vbTab & m2.Text
            For c = 0 To m2.Children.Count - 1
                f.WriteLine progname & vbTab & "#MENU3" & vbTab & a & "/" & b & "/" & c & vbTab & _
                            m2.Children.ElementAt(c).Text
                Err.Clear
            Next
            Err.Clear
        Next
        Err.Clear
    Next
    On Error GoTo 0
End Sub

' dump every grid / list inside the popup window
Sub DumpPopup(progname, tag)
    Dim w
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        f.WriteLine progname & vbTab & tag & vbTab & "NO_POPUP"
        Exit Sub
    End If
    f.WriteLine progname & vbTab & tag & vbTab & "POPUP_TITLE" & vbTab & w.Text
    On Error GoTo 0
    Walk w, progname, tag
End Sub

Sub Walk(root, progname, tag)
    Dim c, k, comp, t, sub_, g2, ord, r, cc, line
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
            sub_ = comp.SubType
            f.WriteLine progname & vbTab & tag & vbTab & "SHELL" & vbTab & _
                        Mid(comp.Id, InStr(comp.Id, "/wnd[1]")) & vbTab & sub_
            If sub_ = "GridView" Then
                Set g2 = comp
                Set ord = g2.ColumnOrder
                For r = 0 To g2.RowCount - 1
                    line = ""
                    For cc = 0 To ord.Count - 1
                        If cc > 0 Then line = line & " | "
                        line = line & ord(cc) & "=" & g2.GetCellValue(r, ord(cc))
                    Next
                    f.WriteLine progname & vbTab & tag & vbTab & "GRIDROW" & vbTab & r & vbTab & line
                    Err.Clear
                    If r > 200 Then Exit For
                Next
            End If
            Err.Clear
        ElseIf t = "GuiTableControl" Then
            f.WriteLine progname & vbTab & tag & vbTab & "TABLECTRL" & vbTab & _
                        Mid(comp.Id, InStr(comp.Id, "/wnd[1]")) & vbTab & comp.RowCount
            Err.Clear
        ElseIf t = "GuiLabel" Or t = "GuiTextField" Or t = "GuiCTextField" Then
            If Trim(comp.Text) <> "" Then
                f.WriteLine progname & vbTab & tag & vbTab & "TXT" & vbTab & _
                            Mid(comp.Id, InStr(comp.Id, "/wnd[1]")) & vbTab & comp.Text
            End If
            Err.Clear
        End If
        Walk comp, progname, tag
    Next
    Err.Clear
End Sub

Sub ClosePopup()
    Dim w
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then
        w.sendVKey 12   ' F12 = cancel, never saves
        WScript.Sleep 900
    End If
    Err.Clear
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then
        w.close
        WScript.Sleep 700
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Sub Reset()
    On Error Resume Next
    ClosePopup
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1100
    ClosePopup
    On Error GoTo 0
End Sub

Sub Probe(progname)
    Dim g, ord, c
    f.WriteLine "############" & vbTab & progname
    Set g = RunQuery(progname)
    If g Is Nothing Then
        f.WriteLine progname & vbTab & "NO_GRID"
        Reset
        Exit Sub
    End If
    On Error Resume Next
    f.WriteLine progname & vbTab & "#TITLE" & vbTab & session.findById("wnd[0]").Text
    f.WriteLine progname & vbTab & "#ROWCOUNT" & vbTab & g.RowCount
    Set ord = g.ColumnOrder
    f.WriteLine progname & vbTab & "#VISIBLECOLS" & vbTab & ord.Count
    For c = 0 To ord.Count - 1
        f.WriteLine progname & vbTab & "#VCOL" & vbTab & c & vbTab & ord(c)
    Next
    Err.Clear
    On Error GoTo 0

    DumpAlvToolbar g, progname
    DumpMenus progname

    ' try to open the change-layout dialog by ALV toolbar function code
    On Error Resume Next
    g.pressToolbarButton "&MB_VARIANT"
    WScript.Sleep 1600
    If Err.Number <> 0 Then
        log.WriteLine progname & " &MB_VARIANT failed: " & Err.Description
        Err.Clear
        g.pressToolbarButton "&OL0"
        WScript.Sleep 1600
        If Err.Number <> 0 Then
            log.WriteLine progname & " &OL0 failed: " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo 0
    DumpPopup progname, "LAYOUT"
    ClosePopup
    Reset
End Sub

Probe "ZCRMR052"
Probe "ZCRMR052S"

f.Close
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "S04_layout_catalog done"
