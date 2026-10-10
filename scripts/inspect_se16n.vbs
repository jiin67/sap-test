' inspect_se16n.vbs : learn the SE16N screen structure so table searches can be scripted
'   READ ONLY. args: [table] (default ZCRM_RECPT)
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i, tbl

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

Dim ARGS
Set ARGS = WScript.Arguments
If ARGS.Count >= 1 Then
    tbl = ARGS(0)
Else
    tbl = "ZCRM_RECPT"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_se16n.txt", True, True)
log.WriteLine "table=" & tbl

n0 = connection.Children.Count
If n0 > 1 Then
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

Sub Status(tag)
    On Error Resume Next
    log.WriteLine "[" & tag & "] prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & _
                  " tcode=" & session.Info.Transaction & _
                  " title=[" & session.findById("wnd[0]").Text & "]" & _
                  " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    Err.Clear
    On Error GoTo 0
End Sub

Sub Tree(comp, depth)
    Dim c, k, line, st
    On Error Resume Next
    line = String(depth * 2, " ") & comp.Type & " | " & Mid(comp.Id, InStr(comp.Id, "/wnd["))
    If comp.Type = "GuiShell" Then
        st = comp.SubType
        If Err.Number = 0 Then line = line & " | sub=" & st
        Err.Clear
    End If
    If comp.Type = "GuiTextField" Or comp.Type = "GuiCTextField" Or comp.Type = "GuiLabel" _
       Or comp.Type = "GuiCheckBox" Or comp.Type = "GuiRadioButton" Then
        line = line & " | text=[" & comp.Text & "]"
    End If
    Err.Clear
    log.WriteLine line
    k = comp.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    For c = 0 To k - 1
        Tree comp.Children.ElementAt(c), depth + 1
    Next
End Sub

Sub DumpGridShell(id, tag)
    Dim g, ord, c, r, line
    On Error Resume Next
    Set g = session.findById(id)
    If Err.Number <> 0 Then
        Err.Clear
        log.WriteLine tag & " : shell not found " & id
        Exit Sub
    End If
    Set ord = g.ColumnOrder
    log.WriteLine tag & " rows=" & g.RowCount & " cols=" & ord.Count
    line = ""
    For c = 0 To ord.Count - 1
        If c > 0 Then line = line & " | "
        line = line & ord(c)
    Next
    log.WriteLine tag & " COLS: " & line
    For r = 0 To g.RowCount - 1
        line = ""
        For c = 0 To ord.Count - 1
            If c > 0 Then line = line & " | "
            line = line & ord(c) & "=" & g.GetCellValue(r, ord(c))
        Next
        log.WriteLine tag & " ROW" & r & ": " & line
        Err.Clear
        If r > 80 Then Exit For
    Next
    Err.Clear
    On Error GoTo 0
End Sub

' ---- open SE16N ----
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2200
If Err.Number <> 0 Then
    log.WriteLine "ERR open SE16N " & Err.Number & " " & Err.Description
    Err.Clear
End If
On Error GoTo 0
Status "SE16N initial"
log.WriteLine "--- wnd[0] tree (initial) ---"
Tree session.findById("wnd[0]/usr"), 0

' ---- enter the table name ----
On Error Resume Next
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = tbl
If Err.Number <> 0 Then
    log.WriteLine "ERR set GD-TAB " & Err.Number & " " & Err.Description
    Err.Clear
End If
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2500
Err.Clear
On Error GoTo 0
Status "after table entered"
log.WriteLine "--- wnd[0] tree (after table) ---"
Tree session.findById("wnd[0]"), 0

' ---- the selection-criteria grid, if any ----
DumpGridShell "wnd[0]/usr/cntlGRID1/shellcont/shell", "SELGRID cntlGRID1"
DumpGridShell "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell", "SELGRID RESULT_LIST"

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
Dim w
Set w = session.findById("wnd[1]")
If Err.Number = 0 Then w.sendVKey 0
Err.Clear
On Error GoTo 0
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_se16n done"
