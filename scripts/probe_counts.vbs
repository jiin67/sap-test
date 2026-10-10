' probe_counts.vbs : row counts for 15-day windows, OLD vs NEW, on KDQ only (READ ONLY)
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

' pick the KDQ connection explicitly (KDD may also be open)
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
    WScript.Echo "GUARD: " & base.Info.SystemName
    WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\probe_counts.txt", True, True)
log.WriteLine "conn sys=" & base.Info.SystemName & " client=" & base.Info.Client & " user=" & base.Info.User
log.WriteLine "sessions=" & connection.Children.Count

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

Sub Open(isNew)
    On Error Resume Next
    If isNew Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 1600
        session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "ZCRMR052S"
        session.findById("wnd[0]/usr/radRS38M-FUNC_EDIT").Select
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        WScript.Sleep 2500
    Else
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 2200
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Sub Q(tag, isNew, org, recv, dfield, dfrom, dto)
    Dim ttl, ok, t0, t1, pname
    If isNew Then
        pname = "NEW"
    Else
        pname = "OLD"
    End If
    Open isNew
    On Error Resume Next
    session.findById("wnd[0]/usr/" & org).Select
    session.findById("wnd[0]/usr/" & recv).Select
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1000
    Err.Clear
    On Error GoTo 0
    ok = SetField("SO_RCDT-LOW", "")
    ok = SetField("SO_RCDT-HIGH", "")
    ok = SetField("SO_PRDRE-LOW", "")
    ok = SetField("SO_PRDRE-HIGH", "")
    ok = SetField(dfield & "-LOW", dfrom)
    ok = SetField(dfield & "-HIGH", dto)
    t0 = Timer
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 1200
    ttl = session.findById("wnd[0]").Text
    Err.Clear
    On Error GoTo 0
    t1 = Timer
    log.WriteLine tag & vbTab & pname & vbTab & dfield & " " & dfrom & "~" & dto & _
                  vbTab & "screen=" & session.Info.ScreenNumber & _
                  vbTab & "count=" & TitleCount(ttl) & _
                  vbTab & "sec=" & FormatNumber(t1 - t0, 1) & _
                  vbTab & "title=[" & ttl & "]"
    On Error Resume Next
    Dim w
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 900
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    On Error GoTo 0
End Sub

' 15-day windows before 2026-07-15
Q "W1 center+recv", False, "radPA_R01", "radPA_CHX", "SO_RCDT", "2026.07.01", "2026.07.15"
Q "W1 center+recv", True,  "radPA_R01", "radPA_CHX", "SO_RCDT", "2026.07.01", "2026.07.15"
Q "W2 center+recv", False, "radPA_R01", "radPA_CHX", "SO_RCDT", "2026.06.01", "2026.06.15"
Q "W2 center+recv", True,  "radPA_R01", "radPA_CHX", "SO_RCDT", "2026.06.01", "2026.06.15"
Q "W1 dealer+recv", False, "radPA_R02", "radPA_CHX", "SO_RCDT", "2026.07.01", "2026.07.15"
Q "W1 dealer+recv", True,  "radPA_R02", "radPA_CHX", "SO_RCDT", "2026.07.01", "2026.07.15"
Q "W1 center+nrcv", False, "radPA_R01", "radPA_NCH", "SO_PRDRE", "2026.07.01", "2026.07.15"
Q "W1 center+nrcv", True,  "radPA_R01", "radPA_NCH", "SO_PRDRE", "2026.07.01", "2026.07.15"

log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "probe_counts done"
