' S09_find_editable.vbs
'   Find rows that the program itself treats as editable - these are the
'   candidates for the save tests (TC-07/08/09) and the lock test (TC-12).
'   Runs the 미접수 query on ZCRM052 and records, per row, whether the input
'   cells are changeable, together with the keys needed to look the row up.
'   READ ONLY : reads cell state only, types nothing, saves nothing.
'   args: <dfrom> <dto> [mode]   mode = NCH (default, 미접수) or CHX (접수)
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

Dim D1, D2, MODE, ARGS
Set ARGS = WScript.Arguments
If ARGS.Count >= 2 Then
    D1 = ARGS(0)
    D2 = ARGS(1)
Else
    D1 = "2026.07.01"
    D2 = "2026.07.15"
End If
If ARGS.Count >= 3 Then
    MODE = UCase(ARGS(2))
Else
    MODE = "NCH"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\raw") Then fso.CreateFolder "C:\sap-test\results\raw"
Set log = fso.CreateTextFile("C:\sap-test\logs\S09_find_editable_" & MODE & ".txt", True, True)
Set f = fso.CreateTextFile("C:\sap-test\results\raw\editable_" & MODE & ".tsv", True, True)
log.WriteLine "range " & D1 & " ~ " & D2 & " mode=" & MODE

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

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2200
If ARGS.Count >= 4 Then
    session.findById("wnd[0]/usr/rad" & UCase(ARGS(3))).Select
Else
    session.findById("wnd[0]/usr/radPA_R01").Select
End If
If MODE = "NCH" Then
    session.findById("wnd[0]/usr/radPA_NCH").Select
Else
    session.findById("wnd[0]/usr/radPA_CHX").Select
End If
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
Err.Clear
On Error GoTo 0

Dim ok
ok = SetField("SO_RCDT-LOW", "")
ok = SetField("SO_RCDT-HIGH", "")
ok = SetField("SO_PRDRE-LOW", "")
ok = SetField("SO_PRDRE-HIGH", "")
If MODE = "NCH" Then
    ok = SetField("SO_PRDRE-LOW", D1)
    ok = SetField("SO_PRDRE-HIGH", D2)
Else
    ok = SetField("SO_RCDT-LOW", D1)
    ok = SetField("SO_RCDT-HIGH", D2)
End If

On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2000
Err.Clear
On Error GoTo 0

Dim g
On Error Resume Next
Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
If Err.Number <> 0 Then
    Err.Clear
    log.WriteLine "no grid"
    f.WriteLine "#GRID" & vbTab & "NOT_FOUND"
    f.Close
    log.Close
    WScript.Quit 10
End If
On Error GoTo 0

' columns we care about for candidate selection
Dim KEYCOLS, EDITCOLS
KEYCOLS = Array("CSLTNO", "REQNO", "RECPTNO", "SRV_QTY", "CF_QTY", "NCF_QTY", _
                "NCFT", "NCF_RSN_VT", "RSPC", "WEARDATS", "KUNNR_SVC", "VKBUR", _
                "MATNR_PRT", "TAGNO1", "PRICETYPET", "PRDREQDATE", "REMARK1", "REMARK2")
EDITCOLS = Array("CF_QTY", "NCF_QTY", "NCF_RSN_VT", "RSPC", "WEARDATS", "REMARK")

Dim nrow, ord, has, c, k
On Error Resume Next
nrow = g.RowCount
Set ord = g.ColumnOrder
On Error GoTo 0
log.WriteLine "rows=" & nrow & " cols=" & ord.Count

Function HasCol(nm)
    Dim j
    HasCol = False
    For j = 0 To ord.Count - 1
        If ord(j) = nm Then
            HasCol = True
            Exit Function
        End If
    Next
End Function

Dim hdr
hdr = "ROW"
For k = 0 To UBound(KEYCOLS)
    If HasCol(KEYCOLS(k)) Then hdr = hdr & vbTab & KEYCOLS(k)
Next
For k = 0 To UBound(EDITCOLS)
    If HasCol(EDITCOLS(k)) Then hdr = hdr & vbTab & "EDIT_" & EDITCOLS(k)
Next
f.WriteLine "#MODE" & vbTab & MODE
f.WriteLine "#RANGE" & vbTab & D1 & vbTab & D2
f.WriteLine "#ROWCOUNT" & vbTab & nrow
f.WriteLine "#DATA"
f.WriteLine hdr

Dim r, line, vis, winTop, maxr, v
maxr = nrow - 1
If maxr > 499 Then maxr = 499
On Error Resume Next
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
    line = r
    For k = 0 To UBound(KEYCOLS)
        If HasCol(KEYCOLS(k)) Then
            v = g.GetCellValue(r, KEYCOLS(k))
            If Err.Number <> 0 Then
                v = ""
                Err.Clear
            End If
            line = line & vbTab & Replace(v, vbTab, " ")
        End If
    Next
    For k = 0 To UBound(EDITCOLS)
        If HasCol(EDITCOLS(k)) Then
            v = g.GetCellChangeable(r, EDITCOLS(k))
            If Err.Number <> 0 Then
                v = "?"
                Err.Clear
            End If
            line = line & vbTab & v
        End If
    Next
    f.WriteLine line
Next
Err.Clear
On Error GoTo 0
log.WriteLine "dumped=" & (maxr + 1)

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
WScript.Echo "S09 " & MODE & " done"
