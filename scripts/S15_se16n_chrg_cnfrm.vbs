' S15_se16n_chrg_cnfrm.vbs : maintain ZCRM_CHRG_CNFRM rows through SE16N edit mode.
'   User-approved (2026-09-11) for scenario group 3 / 4-4 because the RFC rejects every CSLTNO (defect).
'   Only the 4 test CSLTNO may be touched. Inserted rows carry ERNAM=ZTEST0911.
'   args: <tag> <op> [<op> ...]
'     ins:<CSLTNO>:<SPMON YYYY.MM>:<CFDATE1 YYYY.MM.DD>
'     upd:<CSLTNO>:<SPMON YYYY.MM>:<CFDATE1 YYYY.MM.DD>
'     del:<CSLTNO>
'   Saves once at the end and checks the SE16N result popup counts (deleted/inserted/modified)
'   against the plan. Any other popup is logged and cancelled; on any problem nothing is saved.
'   log: logs\S15_<tag>.txt (UTF-16)   exit: 0 saved as planned / 20 aborted (not saved) / 21 counts differ
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, g
Dim TAG, ops(), nOps, nIns, nUpd, nDel, today, allowed

allowed = "|2026070100278|2026070100450|2026070100862|2026070102721|"

If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: S15_se16n_chrg_cnfrm.vbs <tag> ins:..|upd:..|del:.. ..."
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
nOps = WScript.Arguments.Count - 1
ReDim ops(nOps - 1)
nIns = 0 : nUpd = 0 : nDel = 0
For i = 0 To nOps - 1
    ops(i) = Split(WScript.Arguments(i + 1), ":")
    If InStr(allowed, "|" & ops(i)(1) & "|") = 0 Then
        WScript.Echo "CSLTNO not allowed: " & ops(i)(1)
        WScript.Quit 1
    End If
    Select Case ops(i)(0)
        Case "ins" : nIns = nIns + 1
        Case "upd" : nUpd = nUpd + 1
        Case "del" : nDel = nDel + 1
        Case Else
            WScript.Echo "bad op: " & WScript.Arguments(i + 1)
            WScript.Quit 1
    End Select
    If ops(i)(0) <> "del" And UBound(ops(i)) < 3 Then
        WScript.Echo "ins/upd need SPMON and CFDATE1: " & WScript.Arguments(i + 1)
        WScript.Quit 1
    End If
Next
today = Year(Now) & "." & Right("0" & Month(Now), 2) & "." & Right("0" & Day(Now), 2)

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S15_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG & " planned ins=" & nIns & " upd=" & nUpd & " del=" & nDel
For i = 0 To nOps - 1
    log.WriteLine "  op " & Join(ops(i), ":")
Next

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" Then WScript.Quit 99
log.WriteLine "work session=" & session.Id

Function Exists(id)
    Dim o
    On Error Resume Next
    Set o = session.findById(id)
    Exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Function Sbar()
    On Error Resume Next
    Sbar = "[" & session.findById("wnd[0]/sbar").MessageType & "] " & session.findById("wnd[0]/sbar").Text
    Err.Clear
    On Error GoTo 0
End Function

Function CancelPopup(where)
    Dim w, c
    CancelPopup = False
    If Not Exists("wnd[1]") Then Exit Function
    CancelPopup = True
    Set w = session.findById("wnd[1]")
    log.WriteLine "  POPUP (" & where & ") [" & w.Text & "]"
    On Error Resume Next
    For Each c In w.findById("usr").Children
        log.WriteLine "    " & c.Type & " " & c.Name & " [" & c.Text & "]"
    Next
    w.HardCopy "C:\sap-test\results\shots\S15_" & TAG & "_popup_" & where & ".png", 2
    w.sendVKey 12
    WScript.Sleep 1000
    Err.Clear
    On Error GoTo 0
End Function

Sub DumpRows(title)
    Dim r
    log.WriteLine "  ---- " & title & " rows=" & g.RowCount
    On Error Resume Next
    For r = 0 To g.RowCount - 1
        log.WriteLine "    [" & r & "] CSLTNO=" & g.GetCellValue(r, "CSLTNO") & " SPMON=" & g.GetCellValue(r, "SPMON") & _
                      " CFDATE1=" & g.GetCellValue(r, "CFDATE1") & " ERNAM=" & g.GetCellValue(r, "ERNAM") & " ERDAT=" & g.GetCellValue(r, "ERDAT")
        Err.Clear
    Next
    On Error GoTo 0
End Sub

Function FindRow(cs)
    Dim r
    FindRow = -1
    On Error Resume Next
    For r = 0 To g.RowCount - 1
        If Trim(g.GetCellValue(r, "CSLTNO")) = cs Then
            FindRow = r
            Exit For
        End If
    Next
    Err.Clear
    On Error GoTo 0
End Function

Sub Leave(code)
    ' leave without saving anything that is still pending; closing the session discards it
    On Error Resume Next
    log.WriteLine "exit code=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Sleep 1000
    If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
    Err.Clear
    On Error GoTo 0
    WScript.Echo "S15 " & TAG & " exit=" & code
    WScript.Quit code
End Sub

' ---------- open SE16N edit list ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "ZCRM_CHRG_CNFRM"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "50"
session.findById("wnd[0]/usr/chkGD-EDIT").Selected = True
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
If CancelPopup("exec") Then Leave 20
If Not Exists("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
    log.WriteLine "  no edit grid, screen=" & session.Info.ScreenNumber & " sbar=" & Sbar()
    Leave 20
End If
Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
DumpRows "before"

' ---------- apply operations in the grid ----------
Dim r, op
For i = 0 To nOps - 1
    op = ops(i)
    r = FindRow(op(1))
    Select Case op(0)
        Case "ins"
            If r >= 0 Then
                log.WriteLine "  ABORT ins: " & op(1) & " already exists (row " & r & ")"
                Leave 20
            End If
            On Error Resume Next
            g.pressToolbarButton "&LOCAL&APPEND"
            WScript.Sleep 800
            r = g.RowCount - 1
            g.modifyCell r, "CSLTNO", op(1)
            g.modifyCell r, "SPMON", op(2)
            g.modifyCell r, "CFDATE1", op(3)
            g.modifyCell r, "ERNAM", "ZTEST0911"
            g.modifyCell r, "ERDAT", today
            If Err.Number <> 0 Then
                log.WriteLine "  ABORT ins cell write: " & Err.Description
                Err.Clear
                On Error GoTo 0
                Leave 20
            End If
            On Error GoTo 0
        Case "upd"
            If r < 0 Then
                log.WriteLine "  ABORT upd: " & op(1) & " not found"
                Leave 20
            End If
            On Error Resume Next
            g.modifyCell r, "SPMON", op(2)
            g.modifyCell r, "CFDATE1", op(3)
            g.modifyCell r, "AENAM", "ZTEST0911"
            g.modifyCell r, "AEDAT", today
            If Err.Number <> 0 Then
                log.WriteLine "  ABORT upd cell write: " & Err.Description
                Err.Clear
                On Error GoTo 0
                Leave 20
            End If
            On Error GoTo 0
        Case "del"
            If r < 0 Then
                log.WriteLine "  ABORT del: " & op(1) & " not found"
                Leave 20
            End If
            On Error Resume Next
            g.selectedRows = CStr(r)
            g.pressToolbarButton "&LOCAL&DELETE_ROW"
            WScript.Sleep 800
            Err.Clear
            On Error GoTo 0
    End Select
    If CancelPopup("op" & i) Then
        log.WriteLine "  ABORT popup during op " & Join(op, ":")
        Leave 20
    End If
    log.WriteLine "  done op " & Join(op, ":") & " sbar=" & Sbar()
Next
On Error Resume Next
g.pressEnter
WScript.Sleep 1000
Err.Clear
On Error GoTo 0
If CancelPopup("enter") Then
    log.WriteLine "  ABORT popup on check"
    Leave 20
End If
If session.findById("wnd[0]/sbar").MessageType = "E" Then
    log.WriteLine "  ABORT error message before save: " & Sbar()
    Leave 20
End If
DumpRows "before save"
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S15_" & TAG & "_before_save.png", 2

' ---------- save and check the result counts ----------
Dim nD, nI, nM, code
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 2500
Err.Clear
On Error GoTo 0
nD = "" : nI = "" : nM = ""
If Exists("wnd[1]/usr/txtGD_MOD_NR") Then
    nD = Trim(session.findById("wnd[1]/usr/txtGD_DEL_NR").Text)
    nI = Trim(session.findById("wnd[1]/usr/txtGD_INS_NR").Text)
    nM = Trim(session.findById("wnd[1]/usr/txtGD_MOD_NR").Text)
End If
log.WriteLine "  save result: deleted=" & nD & " inserted=" & nI & " modified=" & nM & " sbar=" & Sbar()
CancelPopup "save"
code = 0
If nD <> CStr(nDel) Or nI <> CStr(nIns) Or nM <> CStr(nUpd) Then
    log.WriteLine "  COUNTS DIFFER from plan (del " & nDel & " / ins " & nIns & " / upd " & nUpd & ") - check DB"
    code = 21
End If
If Exists("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    DumpRows "after save"
End If
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S15_" & TAG & "_after_save.png", 2

' ---------- edit mode off, leave ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
Err.Clear
On Error GoTo 0
CancelPopup "leave"
On Error Resume Next
session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
log.WriteLine "edit checkbox final=" & session.findById("wnd[0]/usr/chkGD-EDIT").Selected
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
Err.Clear
On Error GoTo 0
Leave code
