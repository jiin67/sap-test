' inspect_subs_01.vbs : READ-ONLY recon for the subscription (일시불구독) change.
'   1) open each T-code, record program/screen/sbar and dump selection-screen fields
'   2) SE37 : does ZCRMF03_RFC_ZCRM_CHRG_CNFRM exist
'   3) SE16N: selection-field descriptions of ZCRM_CICSVC_NEW / ZCRM_CHRG_CNFRM (no execute)
' No save / no execute of any processing button. Works in its own session.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_subs_01.txt", True, True)

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" And session.Info.SystemName <> "KDD" Then WScript.Quit 99
log.WriteLine "work session=" & session.Id

Sub Go(tc)
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tc
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
End Sub

Sub ClosePopups()
    Dim w, k
    For k = 1 To 3
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        log.WriteLine "  POPUP wnd[1] title=[" & w.Text & "]"
        DumpTree w, 2, 0
        w.sendVKey 12
        WScript.Sleep 800
        Err.Clear
        On Error GoTo 0
    Next
End Sub

Sub DumpTree(comp, depth, lvl)
    Dim c, k, t, s
    On Error Resume Next
    t = comp.Type
    s = String(lvl * 2, " ") & t & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[") + 1)
    s = s & " name=" & comp.Name & " text=[" & comp.Text & "]"
    If t = "GuiShell" Then s = s & " sub=" & comp.SubType
    If t = "GuiRadioButton" Or t = "GuiCheckBox" Then s = s & " sel=" & comp.Selected
    log.WriteLine s
    Err.Clear
    If lvl >= depth Then Exit Sub
    k = comp.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    For c = 0 To k - 1
        DumpTree comp.Children.ElementAt(c), depth, lvl + 1
    Next
End Sub

Sub Status(tag)
    On Error Resume Next
    log.WriteLine tag & " tcode=" & session.Info.Transaction & " prog=" & session.Info.Program & _
        " screen=" & session.Info.ScreenNumber & " title=[" & session.findById("wnd[0]").Text & "]" & _
        " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    Err.Clear
    On Error GoTo 0
End Sub

' ---------- 1) T-codes ----------
Dim tcs, t
tcs = Array("ZCRM058A", "ZCRM058S", "ZCRM059A", "ZCRM059S", "ZCRM060A", "ZCRM060S")
For Each t In tcs
    log.WriteLine "=================== " & t
    Go t
    Status "  after open:"
    ClosePopups
    DumpTree session.findById("wnd[0]/usr"), 3, 0
    ' toolbar buttons on the selection screen
    Dim b
    For i = 0 To 50
        On Error Resume Next
        Set b = session.findById("wnd[0]/tbar[1]/btn[" & i & "]")
        If Err.Number = 0 Then log.WriteLine "  tbar1 btn[" & i & "] text=[" & b.Text & "] tip=[" & b.Tooltip & "]"
        Err.Clear
        On Error GoTo 0
    Next
    Go ""
Next

' ---------- 2) SE37 ----------
log.WriteLine "=================== SE37"
Go "SE37"
On Error Resume Next
DumpTree session.findById("wnd[0]/usr"), 2, 0
session.findById("wnd[0]/usr/ctxtRS38L-NAME").Text = "ZCRMF03_RFC_ZCRM_CHRG_CNFRM"
If Err.Number <> 0 Then log.WriteLine "  ERR name field " & Err.Description : Err.Clear
session.findById("wnd[0]").sendVKey 7
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
Status "  after F7(display):"
ClosePopups
Go ""

' ---------- 3) SE16N selection descriptions ----------
Sub SelDesc(tbl)
    Dim r, pos, maxpos, txt, ed
    log.WriteLine "=================== SE16N " & tbl
    Go "SE16N"
    On Error Resume Next
    Set ed = session.findById("wnd[0]/usr/chkGD-EDIT")
    If Err.Number = 0 Then
        If ed.Selected Then
            log.WriteLine "  WARN edit checkbox was ON - clearing"
            ed.Selected = False
        End If
    End If
    Err.Clear
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = tbl
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
    Status "  after table:"
    ClosePopups
    On Error Resume Next
    maxpos = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VerticalScrollbar.Maximum
    If Err.Number <> 0 Then
        log.WriteLine "  no selection table: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    log.WriteLine "  scroll max=" & maxpos
    pos = 0
    Do
        On Error Resume Next
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VerticalScrollbar.Position = pos
        WScript.Sleep 400
        Err.Clear
        For r = 0 To 40
            txt = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & r & "]").Text
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            log.WriteLine "  pos " & pos & " row " & r & " : [" & txt & "]"
        Next
        On Error GoTo 0
        pos = pos + r
        If r = 0 Then Exit Do
    Loop While pos <= maxpos
End Sub

SelDesc "ZCRM_CICSVC_NEW"
SelDesc "ZCRM_CHRG_CNFRM"
Go ""

log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_subs_01 done"
WScript.Quit 0
