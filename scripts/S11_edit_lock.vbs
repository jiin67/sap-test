' S11_edit_lock.vbs
'   TC-12 : compare the edit restrictions on ONE designated record in both
'   programs. Typically used with a record whose quality inspection is closed
'   (ZCRM_QRST_PRT-ZCLOSE = 'X').
'   READ ONLY : reads cell state only, types nothing, saves nothing.
'   args: <CSLTNO> <RCDT_FROM> <RCDT_TO>
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, f, n0, i
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
    WScript.Echo "GUARD: no KDQ connection"
    WScript.Quit 99
End If
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then
    WScript.Echo "GUARD: " & base.Info.SystemName
    WScript.Quit 99
End If

Dim CS, D1, D2, ARGS
Set ARGS = WScript.Arguments
If ARGS.Count < 3 Then
    WScript.Echo "usage: S11_edit_lock.vbs <CSLTNO> <RCDT_FROM> <RCDT_TO>"
    WScript.Quit 1
End If
CS = ARGS(0)
D1 = ARGS(1)
D2 = ARGS(2)

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder "C:\sap-test\results\shots"
Set log = fso.CreateTextFile("C:\sap-test\logs\S11_edit_lock.txt", True, True)
Set f = fso.CreateTextFile("C:\sap-test\results\raw\TC12_edit_lock.tsv", True, True)
log.WriteLine "CSLTNO=" & CS & " RCDT=" & D1 & "~" & D2
f.WriteLine "#CSLTNO" & vbTab & CS
f.WriteLine "#RCDT" & vbTab & D1 & vbTab & D2

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
log.WriteLine "work session=" & session.Id

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

Sub Reset()
    Dim w, guard
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1500
    Err.Clear
    On Error GoTo 0
    guard = 0
    Do
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        Dim b
        Set b = session.findById("wnd[1]/usr/btnBUTTON_2")
        If Err.Number = 0 Then
            b.press
        Else
            Err.Clear
            w.sendVKey 12
        End If
        WScript.Sleep 1200
        Err.Clear
        On Error GoTo 0
        guard = guard + 1
    Loop While guard < 3
End Sub

Dim EDITCOLS, VALCOLS
EDITCOLS = Array("CF_QTY", "NCF_QTY", "NCF_RSN_VT", "RSPC", "REMARK", "WEARDATS")
VALCOLS = Array("CSLTNO", "REQNO", "RECPTNO", "RECPTDATE", "SRV_QTY", "CF_QTY", _
                "NCF_QTY", "NCFT", "NCF_RSN_VT", "RSPC", "RSPCT", "WEARDATS", _
                "MATNR_PRT", "TAGNO1", "KUNNR_SVC")

Sub Probe(progname)
    Dim tc, ok, g, r, k, line, v
    f.WriteLine "############" & vbTab & progname
    log.WriteLine "=== " & progname
    On Error Resume Next
    If progname = "ZCRMR052" Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
    Else
        tc = NewTcode()
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tc
    End If
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2400
    session.findById("wnd[0]/usr/radPA_R01").Select
    session.findById("wnd[0]/usr/radPA_CHX").Select
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    Err.Clear
    On Error GoTo 0
    ok = SetField("SO_PRDRE-LOW", "")
    ok = SetField("SO_PRDRE-HIGH", "")
    ok = SetField("SO_RCDT-LOW", D1)
    ok = SetField("SO_RCDT-HIGH", D2)
    If CS <> "-" Then ok = SetField("SO_CSLTN-LOW", CS)
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2500
    Err.Clear
    Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
    If Err.Number <> 0 Then
        Err.Clear
        f.WriteLine progname & vbTab & "NO_GRID" & vbTab & session.findById("wnd[0]/sbar").Text
        log.WriteLine "  no grid"
        Reset
        Exit Sub
    End If
    On Error GoTo 0
    f.WriteLine progname & vbTab & "#ROWCOUNT" & vbTab & g.RowCount
    log.WriteLine "  rows=" & g.RowCount
    session.findById("wnd[0]").HardCopy SHOTDIR & "TC12_" & progname & "_result.png", 2

    On Error Resume Next
    For r = 0 To g.RowCount - 1
        line = progname & vbTab & "ROW" & vbTab & r
        For k = 0 To UBound(VALCOLS)
            v = g.GetCellValue(r, VALCOLS(k))
            If Err.Number <> 0 Then
                v = ""
                Err.Clear
            End If
            line = line & vbTab & VALCOLS(k) & "=" & v
        Next
        For k = 0 To UBound(EDITCOLS)
            v = g.GetCellChangeable(r, EDITCOLS(k))
            If Err.Number <> 0 Then
                v = "?"
                Err.Clear
            End If
            line = line & vbTab & "EDIT_" & EDITCOLS(k) & "=" & v
        Next
        f.WriteLine line
        If r > 399 Then Exit For
    Next
    Err.Clear
    On Error GoTo 0
    Reset
End Sub

Probe "ZCRMR052"
Probe "ZCRMR052S"

f.Close
log.WriteLine "done"
log.Close
On Error Resume Next
session.findById("wnd[0]").close
Err.Clear
On Error GoTo 0
WScript.Echo "S11 done"
