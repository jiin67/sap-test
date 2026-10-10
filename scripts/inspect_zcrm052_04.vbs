' inspect_zcrm052_04.vbs : execute ZCRM052 (read-only query) and dump result screen structure
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then
    WScript.Echo "GUARD: " & base.Info.SystemName : WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zcrm052_04.txt", True, True)
log.WriteLine "system=" & base.Info.SystemName & " client=" & base.Info.Client

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

Sub Status(tag)
    log.WriteLine "[" & tag & "] prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & _
                  " title=[" & session.findById("wnd[0]").Text & "]" & _
                  " sbar=[" & session.findById("wnd[0]/sbar").Text & "] mt=" & session.findById("wnd[0]/sbar").MessageType
End Sub

Sub Tree(comp, depth)
    Dim c, k, line, st
    On Error Resume Next
    line = String(depth * 2, " ") & comp.Type & " | " & Mid(comp.Id, InStr(comp.Id, "/wnd["))
    st = comp.SubType
    If Err.Number = 0 And st <> "" Then line = line & " | sub=" & st
    Err.Clear
    log.WriteLine line
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k - 1
        Tree comp.Children.ElementAt(c), depth + 1
    Next
End Sub

Sub RunQuery(dfrom, dto)
    log.WriteLine "=========== ZCRM052  center + received  RCDT " & dfrom & " ~ " & dto & " ==========="
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2500
    session.findById("wnd[0]/usr/radPA_R01").Select
    session.findById("wnd[0]/usr/radPA_CHX").Select
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    session.findById("wnd[0]/usr/ctxtSO_RCDT-LOW").Text = dfrom
    session.findById("wnd[0]/usr/ctxtSO_RCDT-HIGH").Text = dto
    If Err.Number <> 0 Then log.WriteLine "  ERR input " & Err.Number & " " & Err.Description : Err.Clear
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 6000
    If Err.Number <> 0 Then log.WriteLine "  ERR exec " & Err.Number & " " & Err.Description : Err.Clear
    On Error GoTo 0
    Status "AFTER EXEC"
    ' popup?
    On Error Resume Next
    Dim w : Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then
        log.WriteLine "  POPUP [" & w.Text & "]"
        Tree w, 1
    End If
    Err.Clear
    On Error GoTo 0
    log.WriteLine "--- wnd[0] tree ---"
    Tree session.findById("wnd[0]"), 0
End Sub

RunQuery "2026.08.01", "2026.09.09"

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
On Error Resume Next
Dim w2 : Set w2 = session.findById("wnd[1]")
If Err.Number = 0 Then w2.close
Err.Clear
On Error GoTo 0
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_zcrm052_04 done"
