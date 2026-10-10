' S03_detail_and_exit.vbs
'   TC-15 (상세조회) and TC-17 (변경 없이 종료) for both programs.
'   READ ONLY : selects rows and opens the detail popup, never saves.
'   args: <dfrom> <dto>
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i
Dim OUTDIR
OUTDIR = "C:\sap-test\results\raw\"
Dim SHOTDIR
SHOTDIR = "C:\sap-test\results\shots\"

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
Set log = fso.CreateTextFile("C:\sap-test\logs\S03_detail_and_exit.txt", True, True)

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

Dim D1, D2, ARGS
Set ARGS = WScript.Arguments
If ARGS.Count >= 2 Then
    D1 = ARGS(0)
    D2 = ARGS(1)
Else
    D1 = "2026.01.01"
    D2 = "2026.09.09"
End If

Dim f
Set f = fso.CreateTextFile(OUTDIR & "TC15_TC17_detail_exit.tsv", True, True)
f.WriteLine "#RANGE" & vbTab & D1 & vbTab & D2

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

Function OpenAndRun(progname)
    Dim g, ok
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
    ok = SetField("SO_RCDT-LOW", D1)
    ok = SetField("SO_RCDT-HIGH", D2)
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 9000
    Err.Clear
    Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
    If Err.Number <> 0 Then
        Err.Clear
        Set OpenAndRun = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set OpenAndRun = g
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


Sub Shot(nm)
    Dim pth
    pth = SHOTDIR & nm & ".png"
    On Error Resume Next
    session.findById("wnd[0]").HardCopy pth, 2
    Err.Clear
    On Error GoTo 0
End Sub

Sub ShotPopup(nm)
    Dim pth, w
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    pth = SHOTDIR & nm & ".png"
    w.HardCopy pth, 2
    If Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[0]").HardCopy pth, 2
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub DumpPopup(tag)
    Dim w, g2, r, c, ord, line
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        f.WriteLine tag & vbTab & "NO_POPUP" & vbTab & "sbar=[" & session.findById("wnd[0]/sbar").Text & _
                    "] mt=" & session.findById("wnd[0]/sbar").MessageType
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    f.WriteLine tag & vbTab & "POPUP" & vbTab & "title=[" & w.Text & "]"
    ' popup ALV?
    Set g2 = session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell")
    If Err.Number <> 0 Then
        Err.Clear
        ' fall back: dump labels / text fields in the popup
        DumpPopupTree w, 1, tag
    Else
        Set ord = g2.ColumnOrder
        line = ""
        For c = 0 To ord.Count - 1
            If c > 0 Then line = line & vbTab
            line = line & ord(c)
        Next
        f.WriteLine tag & vbTab & "POPUPCOLS" & vbTab & line
        For r = 0 To g2.RowCount - 1
            line = ""
            For c = 0 To ord.Count - 1
                If c > 0 Then line = line & vbTab
                line = line & g2.GetCellValue(r, ord(c))
            Next
            f.WriteLine tag & vbTab & "POPUPROW" & vbTab & r & vbTab & line
            If r > 20 Then Exit For
        Next
    End If
    Err.Clear
    w.close
    WScript.Sleep 800
    Err.Clear
    On Error GoTo 0
End Sub

Sub DumpPopupTree(root, depth, tag)
    Dim c, k, comp, t
    On Error Resume Next
    k = root.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For c = 0 To k - 1
        Set comp = root.Children.ElementAt(c)
        t = comp.Type
        If t = "GuiLabel" Or t = "GuiTextField" Or t = "GuiCTextField" Or t = "GuiShell" Then
            f.WriteLine tag & vbTab & "POPUPNODE" & vbTab & t & vbTab & _
                        Mid(comp.Id, InStr(comp.Id, "/wnd[1]")) & vbTab & "[" & comp.Text & "]"
        End If
        DumpPopupTree comp, depth + 1, tag
    Next
    Err.Clear
End Sub

Sub Reset()
    On Error Resume Next
    Dim w
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1300
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    On Error GoTo 0
End Sub

Sub TC15(progname)
    Dim g
    f.WriteLine "########" & vbTab & "TC15" & vbTab & progname
    Set g = OpenAndRun(progname)
    If g Is Nothing Then
        f.WriteLine "TC15 " & progname & vbTab & "NO_GRID"
        Reset
        Exit Sub
    End If
    f.WriteLine "TC15 " & progname & vbTab & "ROWCOUNT" & vbTab & g.RowCount
    If g.RowCount = 0 Then
        f.WriteLine "TC15 " & progname & vbTab & "SKIP_NO_DATA"
        Reset
        Exit Sub
    End If

    On Error Resume Next
    ' (a) no row selected
    g.selectedRows = ""
    session.findById("wnd[0]/tbar[1]/btn[23]").press
    WScript.Sleep 2000
    Err.Clear
    On Error GoTo 0
    ShotPopup "TC15_" & progname & "_nosel"
    Shot "TC15_" & progname & "_nosel_main"
    DumpPopup "TC15 " & progname & " NOSEL"

    ' (b) one row selected
    On Error Resume Next
    g.setCurrentCell 0, g.ColumnOrder(0)
    g.selectedRows = "0"
    session.findById("wnd[0]/tbar[1]/btn[23]").press
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
    ShotPopup "TC15_" & progname & "_onesel"
    DumpPopup "TC15 " & progname & " ONESEL"

    ' (c) two rows selected
    If g.RowCount >= 2 Then
        On Error Resume Next
        g.selectedRows = "0,1"
        session.findById("wnd[0]/tbar[1]/btn[23]").press
        WScript.Sleep 2000
        Err.Clear
        On Error GoTo 0
        ShotPopup "TC15_" & progname & "_twosel"
        Shot "TC15_" & progname & "_twosel_main"
        DumpPopup "TC15 " & progname & " TWOSEL"
    End If
    Reset
End Sub

Sub TC17(progname)
    Dim g
    f.WriteLine "########" & vbTab & "TC17" & vbTab & progname
    Set g = OpenAndRun(progname)
    If g Is Nothing Then
        f.WriteLine "TC17 " & progname & vbTab & "NO_GRID"
        Reset
        Exit Sub
    End If
    ' no change at all, just go back (F3)
    On Error Resume Next
    session.findById("wnd[0]").sendVKey 3
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
    ShotPopup "TC17_" & progname & "_back_popup"
    Shot "TC17_" & progname & "_back"
    DumpPopup "TC17 " & progname & " BACK"
    On Error Resume Next
    f.WriteLine "TC17 " & progname & vbTab & "AFTER_BACK" & vbTab & _
                "prog=" & session.Info.Program & " screen=" & session.Info.ScreenNumber & _
                " title=[" & session.findById("wnd[0]").Text & "]"
    Err.Clear
    On Error GoTo 0
    Reset
End Sub

TC15 "ZCRMR052"
TC15 "ZCRMR052S"
TC17 "ZCRMR052"
TC17 "ZCRMR052S"

f.Close
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "S03_detail_and_exit done"
