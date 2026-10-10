' inspect_zcrm052_03.vbs : selection-screen field/enable comparison, ZCRM052 vs ZCRMR052S
' READ ONLY. Does not execute the report (no F8 on the selection screen).
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then
    WScript.Echo "GUARD: system " & base.Info.SystemName & " not allowed"
    WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zcrm052_03.txt", True, True)
log.WriteLine "system=" & base.Info.SystemName & " client=" & base.Info.Client & " user=" & base.Info.User

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
    log.WriteLine "  [" & tag & "] prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & _
                  " tcode=" & session.Info.Transaction & " title=[" & session.findById("wnd[0]").Text & "]" & _
                  " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
End Sub

Sub DumpSel(root, depth)
    Dim c, k, comp, t, sid, line
    On Error Resume Next
    k = root.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    For c = 0 To k - 1
        Set comp = root.Children.ElementAt(c)
        t = comp.Type
        sid = comp.Id
        sid = Mid(sid, InStr(sid, "/wnd["))
        If t = "GuiTextField" Or t = "GuiCTextField" Or t = "GuiRadioButton" Or _
           t = "GuiCheckBox" Or t = "GuiComboBox" Or t = "GuiLabel" Then
            line = t & " | " & sid & " | text=[" & comp.Text & "]"
            line = line & " | chg=" & comp.Changeable
            If t = "GuiRadioButton" Then line = line & " | sel=" & comp.Selected
            Err.Clear
            log.WriteLine line
        End If
        DumpSel comp, depth + 1
    Next
    Err.Clear
End Sub

Sub Snapshot(tag)
    log.WriteLine "--------- " & tag & " ---------"
    Status tag
    On Error Resume Next
    DumpSel session.findById("wnd[0]/usr"), 0
    Err.Clear
    On Error GoTo 0
End Sub

Sub ClosePopups()
    Dim w, g
    g = 0
    Do
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number <> 0 Then Err.Clear : Exit Do
        On Error GoTo 0
        log.WriteLine "  [popup] [" & w.Text & "]"
        DumpSel w, 0
        w.close
        WScript.Sleep 600
        g = g + 1
    Loop While g < 3
End Sub

Sub PressRadio(rid)
    On Error Resume Next
    session.findById("wnd[0]/usr/" & rid).Select
    If Err.Number <> 0 Then
        log.WriteLine "  ERR select " & rid & " : " & Err.Number & " " & Err.Description
        Err.Clear
        Exit Sub
    End If
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    Err.Clear
    On Error GoTo 0
End Sub

Sub Survey(label)
    Dim j, b
    log.WriteLine "###################### " & label & " ######################"
    Snapshot "INITIAL"
    PressRadio "radPA_R01" : Snapshot "PA_R01 center"
    PressRadio "radPA_R02" : Snapshot "PA_R02 dealer"
    PressRadio "radPA_R01" : Snapshot "back PA_R01"
    PressRadio "radPA_CHX" : Snapshot "PA_CHX received"
    PressRadio "radPA_NCH" : Snapshot "PA_NCH not-received"
    PressRadio "radPA_CHX" : Snapshot "back PA_CHX"
    log.WriteLine "--- toolbar tbar[1] ---"
    On Error Resume Next
    For j = 0 To 20
        Set b = session.findById("wnd[0]/tbar[1]/btn[" & j & "]")
        If Err.Number = 0 Then log.WriteLine "  tbar[1]/btn[" & j & "] text=[" & b.Text & "] tip=[" & b.Tooltip & "]"
        Err.Clear
    Next
    On Error GoTo 0
End Sub

' ---------- old program via tcode ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2500
If Err.Number <> 0 Then log.WriteLine "ERR open ZCRM052 " & Err.Description : Err.Clear
On Error GoTo 0
ClosePopups
Survey "OLD ZCRM052 = ZCRMR052"

' ---------- new program via SE38 execute ----------
log.WriteLine ""
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "ZCRMR052S"
session.findById("wnd[0]/usr/radRS38M-FUNC_EDIT").Select
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 3000
If Err.Number <> 0 Then log.WriteLine "ERR open ZCRMR052S " & Err.Number & " " & Err.Description : Err.Clear
On Error GoTo 0
ClosePopups
Survey "NEW ZCRMR052S"

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 900
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_zcrm052_03 done"
