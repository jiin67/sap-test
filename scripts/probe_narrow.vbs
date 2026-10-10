' probe_narrow.vbs : find a SMALL date window that still returns rows (READ ONLY)
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then
    WScript.Echo "GUARD " & base.Info.SystemName
    WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\probe_narrow.txt", True, True)
log.WriteLine "sessions at start=" & connection.Children.Count

' close any leftover extra sessions except ses[0]
Do While connection.Children.Count > 1
    On Error Resume Next
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    session.findById("wnd[0]").close
    WScript.Sleep 1200
    If Err.Number <> 0 Then
        log.WriteLine "close err " & Err.Description
        Err.Clear
        Exit Do
    End If
    On Error GoTo 0
Loop
log.WriteLine "sessions after cleanup=" & connection.Children.Count

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
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
        TitleCount = Trim(Mid(t, a + 1, b - a - 1))
    Else
        TitleCount = ""
    End If
End Function

Sub Q(tag, org, recv, dfield, dfrom, dto)
    Dim ttl, ok, t0, t1
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2000
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
    WScript.Sleep 1500
    ttl = session.findById("wnd[0]").Text
    Err.Clear
    On Error GoTo 0
    t1 = Timer
    log.WriteLine tag & " | " & dfield & " " & dfrom & "~" & dto & _
                  " | screen=" & session.Info.ScreenNumber & _
                  " | count=" & TitleCount(ttl) & " | title=[" & ttl & "]" & _
                  " | sec=" & FormatNumber(t1 - t0, 1)
    On Error Resume Next
    Dim w
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    On Error GoTo 0
End Sub

Q "d0901", "radPA_R01", "radPA_CHX", "SO_RCDT", "2026.09.01", "2026.09.01"
Q "d0902", "radPA_R01", "radPA_CHX", "SO_RCDT", "2026.09.02", "2026.09.02"
Q "d0903", "radPA_R01", "radPA_CHX", "SO_RCDT", "2026.09.03", "2026.09.03"
Q "d0826", "radPA_R01", "radPA_CHX", "SO_RCDT", "2026.08.26", "2026.08.26"
Q "dealer_d0901", "radPA_R02", "radPA_CHX", "SO_RCDT", "2026.09.01", "2026.09.01"
Q "nch_d0901", "radPA_R01", "radPA_NCH", "SO_PRDRE", "2026.09.01", "2026.09.01"

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "probe_narrow done"
