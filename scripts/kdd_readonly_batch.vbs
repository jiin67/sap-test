' kdd_readonly_batch.vbs : READ-ONLY. close leftover KDD sessions, USR05 FBGENDAT users, SE37 test data dir + source display shots.
Option Explicit
Const LOGNAME = "kdd_readonly_batch"
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For i = 0 To application.Children.Count - 1
    Set base = application.Children.ElementAt(i).Children.ElementAt(0)
    If base.Info.SystemName = "KDD" Then Set connection = application.Children.ElementAt(i) : Exit For
Next
If connection Is Nothing Then WScript.Echo "GUARD no KDD" : WScript.Quit 99
Set base = connection.Children.ElementAt(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\" & LOGNAME & ".txt", True, True)
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDD" Then WScript.Quit 99
log.WriteLine "sys=" & session.Info.SystemName & " client=" & session.Info.Client & " user=" & session.Info.User & " session=" & session.Id & " at=" & Now

Function Exists(id)
    Dim o
    On Error Resume Next
    Set o = session.findById(id)
    Exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
Function T(id)
    On Error Resume Next
    T = ""
    T = session.findById(id).Text
    Err.Clear
    On Error GoTo 0
End Function
Sub Status(t)
    On Error Resume Next
    log.WriteLine t & " prog=" & session.Info.Program & " screen=" & session.Info.ScreenNumber & _
        " title=[" & session.findById("wnd[0]").Text & "] sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    Err.Clear
    On Error GoTo 0
End Sub
Sub DumpTree(comp, depth, lvl)
    Dim c, k, t, s
    On Error Resume Next
    t = comp.Type
    s = String(lvl * 2, " ") & t & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[") + 1) & " text=[" & comp.Text & "]"
    If t = "GuiShell" Then s = s & " sub=" & comp.SubType
    s = s & " tip=[" & comp.Tooltip & "] chg=" & comp.Changeable
    log.WriteLine s
    Err.Clear
    If lvl >= depth Then Exit Sub
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k - 1
        DumpTree comp.Children.ElementAt(c), depth, lvl + 1
    Next
End Sub
Sub Shot(name)
    On Error Resume Next
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\" & name & ".png", 2
    Err.Clear
    On Error GoTo 0
End Sub
Sub Tcode(t)
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1800
    Err.Clear
    On Error GoTo 0
End Sub
Sub Finish(code)
    On Error Resume Next
    Tcode "/n"
    If Exists("wnd[1]") Then log.WriteLine "popup on exit [" & T("wnd[1]") & "]" : session.findById("wnd[1]").sendVKey 12 : WScript.Sleep 800
    log.WriteLine "exit=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Echo LOGNAME & " exit=" & code
    WScript.Quit code
End Sub
' ---- close leftover script sessions (index >= 1, not ours)
Dim k, s
For k = connection.Children.Count - 2 To 1 Step -1
    Set s = connection.Children.ElementAt(k)
    If s.Id <> session.Id Then
        log.WriteLine "closing leftover " & s.Id & " tcode=" & s.Info.Transaction
        s.findById("wnd[0]").close
        WScript.Sleep 1500
    End If
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
log.WriteLine "work session now " & session.Id & " count=" & connection.Children.Count

' ---- 1. SE16N USR05 PARID = FBGENDAT
Tcode "/nSE16N"
Status "SE16N:"
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "USR05"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "200"
If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then
    log.WriteLine "WARN GD-EDIT was ON -> switching OFF"
    session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
End If
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
Dim tbl, r, d, found
found = False
Set tbl = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC")
For r = 0 To tbl.RowCount - 1
    d = ""
    On Error Resume Next
    d = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & r & "]").Text
    On Error GoTo 0
    log.WriteLine "  cond row " & r & " [" & d & "]"
    If InStr(UCase(d), "ID") > 0 And Not found Then
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & r & "]").Text = "FBGENDAT"
        found = True
        log.WriteLine "  -> set FBGENDAT on row " & r
    End If
Next
If found Then
    session.findById("wnd[0]").sendVKey 8
    WScript.Sleep 3000
    Status "USR05 result:"
    If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : session.findById("wnd[1]").sendVKey 12
    Dim g, cols, c, line
    On Error Resume Next
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number = 0 Then
        Set cols = g.ColumnOrder
        line = ""
        For c = 0 To cols.Count - 1 : line = line & cols(c) & vbTab : Next
        log.WriteLine "  rows=" & g.RowCount & " | " & line
        For r = 0 To g.RowCount - 1
            line = ""
            For c = 0 To cols.Count - 1 : line = line & g.GetCellValue(r, cols(c)) & vbTab : Next
            log.WriteLine "  " & r & ": " & line
        Next
    Else
        log.WriteLine "  no result grid err=" & Err.Description
    End If
    Err.Clear
    On Error GoTo 0
    Shot "USR05_fbgendat"
End If

' ---- 2. SE37 test data directory (our FM)
Tcode "/nSE37"
session.findById("wnd[0]/usr/ctxtRS38L-NAME").Text = "ZCRMF03_RFC_ZCRM_CHRG_CNFRM"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 2500
Status "SE37 test screen:"
If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : session.findById("wnd[1]").sendVKey 12 : WScript.Sleep 800
If Exists("wnd[0]/tbar[1]/btn[42]") Then
    Shot "FBG_probe_se37_test"
    session.findById("wnd[0]/tbar[1]/btn[42]").press
    WScript.Sleep 2500
    Status "test data dir:"
    If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : DumpTree session.findById("wnd[1]"), 4, 0
    DumpTree session.findById("wnd[0]/usr"), 4, 0
    log.WriteLine "---- tbar[1]"
    DumpTree session.findById("wnd[0]/tbar[1]"), 2, 0
    Shot "FBG_probe_tdd_before"
Else
    log.WriteLine "  no btn[42] on test screen"
End If

' ---- 3. SE37 display source
Tcode "/nSE37"
session.findById("wnd[0]/usr/ctxtRS38L-NAME").Text = "ZCRMF03_RFC_ZCRM_CHRG_CNFRM"
session.findById("wnd[0]").sendVKey 7
WScript.Sleep 3000
Status "SE37 display:"
If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : session.findById("wnd[1]").sendVKey 0 : WScript.Sleep 800
DumpTree session.findById("wnd[0]/usr"), 2, 0
Shot "FBG_probe_se37_source"

' ---- 4. SE38 display LZCRMG03TOP
Tcode "/nSE38"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "LZCRMG03TOP"
session.findById("wnd[0]/usr/radRS38M-FUNC_EDIT").Select
session.findById("wnd[0]/usr/btnSHOP").press
WScript.Sleep 3000
Status "SE38 TOP display:"
If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : session.findById("wnd[1]").sendVKey 0 : WScript.Sleep 800
Shot "FBG_probe_top_source"
Finish 0
