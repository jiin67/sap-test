' probe_rsn_text.vbs : read how the 부적합사유 column is displayed on screen,
'   using an existing record that already has Z02 stored. READ ONLY.
'   args: <CSLTNO> <RCDT>
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

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

Dim CS, RD, ARGS
Set ARGS = WScript.Arguments
CS = ARGS(0)
RD = ARGS(1)

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\probe_rsn_text.txt", True, True)

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
log.WriteLine "work session=" & session.Id & " CSLTNO=" & CS & " RCDT=" & RD

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

Dim ok
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2200
session.findById("wnd[0]/usr/radPA_R01").Select
session.findById("wnd[0]/usr/radPA_CHX").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
Err.Clear
On Error GoTo 0
ok = SetField("SO_RCDT-LOW", RD)
ok = SetField("SO_RCDT-HIGH", RD)
ok = SetField("SO_CSLTN-LOW", CS)

On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2500
Err.Clear
On Error GoTo 0

Dim g, r
On Error Resume Next
Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
If Err.Number <> 0 Then
    Err.Clear
    log.WriteLine "no grid (maybe 0 rows) sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    log.Close
    WScript.Quit 10
End If
log.WriteLine "rows=" & g.RowCount
For r = 0 To g.RowCount - 1
    log.WriteLine "row" & r & _
        "  NCFT=[" & g.GetCellValue(r, "NCFT") & "]" & _
        "  NCF_RSN_VT=[" & g.GetCellValue(r, "NCF_RSN_VT") & "]" & _
        "  NCF_RSN_V=[" & g.GetCellValue(r, "NCF_RSN_V") & "]" & _
        "  CF_QTY=[" & g.GetCellValue(r, "CF_QTY") & "]" & _
        "  NCF_QTY=[" & g.GetCellValue(r, "NCF_QTY") & "]" & _
        "  RSPC=[" & g.GetCellValue(r, "RSPC") & "]" & _
        "  RSPCT=[" & g.GetCellValue(r, "RSPCT") & "]"
    Err.Clear
    If r > 5 Then Exit For
Next
Err.Clear
On Error GoTo 0

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "probe_rsn_text done"
