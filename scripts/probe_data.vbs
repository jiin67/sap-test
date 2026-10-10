' probe_data.vbs : find selection conditions that return rows (READ ONLY)
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\probe_data.txt", True, True)

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

Function GridRows()
    Dim g
    On Error Resume Next
    Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
    If Err.Number <> 0 Then Err.Clear : GridRows = -1 : Exit Function
    GridRows = g.RowCount
    If Err.Number <> 0 Then Err.Clear : GridRows = -2
    On Error GoTo 0
End Function

Sub Q(tag, org, recv, dfield, dfrom, dto)
    Dim rows, ttl, sb
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2200
    session.findById("wnd[0]/usr/" & org).Select
    session.findById("wnd[0]/usr/" & recv).Select
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    ' clear both date ranges first when present
    session.findById("wnd[0]/usr/ctxtSO_RCDT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSO_RCDT-HIGH").Text = ""
    Err.Clear
    session.findById("wnd[0]/usr/ctxtSO_PRDRE-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSO_PRDRE-HIGH").Text = ""
    Err.Clear
    If dfield <> "" Then
        session.findById("wnd[0]/usr/ctxt" & dfield & "-LOW").Text = dfrom
        session.findById("wnd[0]/usr/ctxt" & dfield & "-HIGH").Text = dto
        If Err.Number <> 0 Then log.WriteLine tag & " : ERR set " & dfield & " " & Err.Description : Err.Clear
    End If
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 8000
    If Err.Number <> 0 Then log.WriteLine tag & " : ERR exec " & Err.Number & " " & Err.Description : Err.Clear
    ttl = session.findById("wnd[0]").Text
    sb = session.findById("wnd[0]/sbar").Text
    On Error GoTo 0
    rows = GridRows()
    log.WriteLine tag & " | screen=" & session.Info.ScreenNumber & " | title=[" & ttl & "] | rows=" & rows & " | sbar=[" & sb & "]"
    ' clear popups
    On Error Resume Next
    Dim w : Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then w.sendVKey 0
    Err.Clear
    On Error GoTo 0
End Sub

Q "A center+received  RCDT 2026",  "radPA_R01", "radPA_CHX", "SO_RCDT",  "2026.01.01", "2026.09.09"
Q "B center+received  RCDT 2025",  "radPA_R01", "radPA_CHX", "SO_RCDT",  "2025.01.01", "2025.12.31"
Q "C dealer+received  RCDT 2026",  "radPA_R02", "radPA_CHX", "SO_RCDT",  "2026.01.01", "2026.09.09"
Q "D dealer+received  RCDT 2025",  "radPA_R02", "radPA_CHX", "SO_RCDT",  "2025.01.01", "2025.12.31"
Q "E center+notrecv   PRDRE 2026", "radPA_R01", "radPA_NCH", "SO_PRDRE", "2026.01.01", "2026.09.09"
Q "F center+notrecv   PRDRE 2025", "radPA_R01", "radPA_NCH", "SO_PRDRE", "2025.01.01", "2025.12.31"
Q "G dealer+notrecv   PRDRE 2026", "radPA_R02", "radPA_NCH", "SO_PRDRE", "2026.01.01", "2026.09.09"
Q "H dealer+notrecv   PRDRE 2025", "radPA_R02", "radPA_NCH", "SO_PRDRE", "2025.01.01", "2025.12.31"

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "probe_data done"
