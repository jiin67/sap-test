' S06_evidence.vbs
'   Capture screenshots (PNG) of the selection screen and the result screen for
'   every comparison case, on BOTH programs, for the test report.
'   READ ONLY : runs queries and captures the window, never saves.
'   args: <dfrom> <dto> [caseFilter]
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i
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

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder "C:\sap-test\results\shots"
Set log = fso.CreateTextFile("C:\sap-test\logs\S06_evidence.txt", True, True)

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

Dim D1, D2, ONLY, ARGS
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
log.WriteLine "range " & D1 & " ~ " & D2 & " filter=" & ONLY

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

Sub Shot(name)
    Dim p
    p = SHOTDIR & name & ".png"
    On Error Resume Next
    session.findById("wnd[0]").HardCopy p, 2
    If Err.Number <> 0 Then
        log.WriteLine "  SHOT ERR " & name & " : " & Err.Description
        Err.Clear
    Else
        log.WriteLine "  shot " & name & ".png"
    End If
    On Error GoTo 0
End Sub

Sub ShotPopup(name)
    Dim p, w
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    p = SHOTDIR & name & ".png"
    w.HardCopy p, 2
    If Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[0]").HardCopy p, 2
        Err.Clear
    End If
    log.WriteLine "  shot(popup) " & name & ".png"
    On Error GoTo 0
End Sub

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

Sub ClearDates()
    Dim r
    r = SetField("SO_RCDT-LOW", "")
    r = SetField("SO_RCDT-HIGH", "")
    r = SetField("SO_PRDRE-LOW", "")
    r = SetField("SO_PRDRE-HIGH", "")
End Sub

Sub ApplySpec(spec)
    Dim parts, p, kv, r
    If spec = "" Then Exit Sub
    parts = Split(spec, ";")
    For Each p In parts
        If Trim(p) <> "" Then
            kv = Split(p, "=")
            If UBound(kv) >= 1 Then r = SetField(Trim(kv(0)), Trim(kv(1)))
        End If
    Next
End Sub

Sub Open(isNew)
    Dim tc
    On Error Resume Next
    If isNew Then
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
    Else
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 2200
    End If
    Err.Clear
    On Error GoTo 0
End Sub

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
    Dim tag, ttl
    If isNew Then
        tag = caseid & "_NEW"
    Else
        tag = caseid & "_OLD"
    End If
    log.WriteLine "== " & tag
    Open isNew
    On Error Resume Next
    session.findById("wnd[0]/usr/" & org).Select
    session.findById("wnd[0]/usr/" & recv).Select
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1100
    Err.Clear
    On Error GoTo 0
    ClearDates
    ApplySpec spec
    Shot tag & "_1_selection"
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 1500
    ttl = session.findById("wnd[0]").Text
    Err.Clear
    On Error GoTo 0
    Shot tag & "_2_result"
    log.WriteLine "   title=[" & ttl & "]"
    Reset
End Sub

Sub RunCase(caseid, org, recv, spec)
    If ONLY <> "ALL" And InStr(caseid, ONLY) <> 1 Then Exit Sub
    log.WriteLine "########### " & caseid
    RunOne caseid, False, org, recv, spec
    RunOne caseid, True, org, recv, spec
End Sub

RunCase "TC01", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2
RunCase "TC02", "radPA_R02", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2
RunCase "TC03", "radPA_R01", "radPA_NCH", "SO_PRDRE-LOW=" & D1 & ";SO_PRDRE-HIGH=" & D2
RunCase "TC04Y", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2 & ";SO_PRTY-LOW=Y"
RunCase "TC04N", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2 & ";SO_PRTY-LOW=N"
RunCase "TC05O", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2 & ";SO_MULTK-LOW=O"
RunCase "TC05X", "radPA_R01", "radPA_CHX", "SO_RCDT-LOW=" & D1 & ";SO_RCDT-HIGH=" & D2 & ";SO_MULTK-LOW=X"

log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "S06_evidence done"
