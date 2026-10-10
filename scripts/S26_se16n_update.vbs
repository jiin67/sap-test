' S26_se16n_update.vbs : UPDATE fields of ONE existing test row through SE16N edit mode (based on S19).
'   2026-09-30: fix test rows of 2026093090011~020 whose common overrides were lost (PowerShell $C/$c name clash).
'   User-approved: test CSLTNO 2026091490001~003 only, tables ZCRM_CICSVC / ZCRM_CICCOST_M / ZCRM_CICCOST_SUB / ZCRM_CICCOST_ADD.
'   The template row is NEVER opened in edit mode: its values come from a read-only S12 TSV dump.
'   args: <tag> <op> [<op> ...]
'     ins|<TABLE>|<NEW CSLTNO>|<template tsv or ->|F=v;F=v;...
'       - opens SE16N edit mode with EXACT 상담번호=<NEW>; the grid must be empty (row must not exist)
'       - appends one row, fills every template column (icons @..@ skipped), then the overrides (F=v; empty v = clear)
'       - save result popup must show inserted=1 deleted=0 modified=0
'   log: logs\S19_<tag>.txt (UTF-16)   shots: results\shots\S19_<tag>_<table>_<csltno>_*.png
'   exit: 0 ok / 1 bad args / 20 some op failed.   File must be saved as CP949.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, pp
Dim TAG, nFail, allowedCs, allowedTb

allowedCs = "|2026091490001|2026091490002|2026091490003|"
allowedTb = "|ZCRM_CICSVC|ZCRM_CICSVC_NEW|ZCRM_CICCOST_M|ZCRM_CICCOST_SUB|ZCRM_CICCOST_ADD|ZCRM_REQ|ZCRM_CHRG_CNFRM|"
If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: S26_se16n_update.vbs <tag> upd|TABLE|CSLTNO|-|F=v;F=v ..."
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
For i = 1 To WScript.Arguments.Count - 1
    pp = Split(WScript.Arguments(i), "|")
    If UBound(pp) < 4 Then
        WScript.Echo "bad op: " & WScript.Arguments(i)
        WScript.Quit 1
    End If
    If pp(0) <> "upd" Or InStr(allowedTb, "|" & pp(1) & "|") = 0 Or (InStr(allowedCs, "|" & pp(2) & "|") = 0 And Left(pp(2), 11) <> "20260930900" And Left(pp(2), 9) <> "202610029") Then
        WScript.Echo "op not allowed: " & WScript.Arguments(i)
        WScript.Quit 1
    End If
Next

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For i = 0 To application.Children.Count - 1
    If application.Children.ElementAt(i).Children.ElementAt(0).Info.SystemName = "KDQ" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then WScript.Quit 99
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S26_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG
If connection.Children.Count >= 6 Then
    log.WriteLine "ABORT: session limit"
    log.Close
    WScript.Quit 14
End If
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" Then WScript.Quit 99
If session.Info.Transaction <> "SESSION_MANAGER" And session.Info.Transaction <> "SMEN" Then
    log.WriteLine "ABORT: picked session is busy (" & session.Info.Transaction & ")"
    log.Close
    WScript.Quit 14
End If
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

Sub Go(tc)
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tc
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2000
    Err.Clear
    On Error GoTo 0
End Sub

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
    w.HardCopy "C:\sap-test\results\shots\S19_" & TAG & "_popup_" & where & ".png", 2
    w.sendVKey 12
    WScript.Sleep 1000
    Err.Clear
    On Error GoTo 0
End Function

Sub EditOff()
    On Error Resume Next
    session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    Err.Clear
    On Error GoTo 0
End Sub

Function SetSelByDesc(desc, lowv)
    Dim r, txt, pos, maxpos
    SetSelByDesc = False
    On Error Resume Next
    maxpos = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VerticalScrollbar.Maximum
    Err.Clear
    pos = 0
    Do
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VerticalScrollbar.Position = pos
        WScript.Sleep 400
        Err.Clear
        For r = 0 To 40
            txt = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & r & "]").Text
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            If Trim(txt) = desc Then
                session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & r & "]").Text = lowv
                If Err.Number = 0 Then SetSelByDesc = True
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
        Next
        pos = pos + r
        If r = 0 Then Exit Do
    Loop While pos <= maxpos
    On Error GoTo 0
End Function

' read header + first data row of an S12 TSV (UTF-16)
Function ReadTpl(path, ByRef names, ByRef vals)
    Dim f, ln, hdr, dat, stage
    ReadTpl = False
    If path = "-" Then Exit Function
    If Not fso.FileExists(path) Then
        log.WriteLine "  template tsv missing: " & path
        Exit Function
    End If
    Set f = fso.OpenTextFile(path, 1, False, -1)
    stage = 0
    Do While Not f.AtEndOfStream
        ln = f.ReadLine
        If stage = 0 Then
            If ln = "#DATA" Then stage = 1
        ElseIf stage = 1 Then
            hdr = ln
            stage = 2
        ElseIf stage = 2 Then
            dat = ln
            stage = 3
            Exit Do
        End If
    Loop
    f.Close
    If stage <> 3 Then
        log.WriteLine "  template tsv has no data row: " & path
        Exit Function
    End If
    names = Split(hdr, vbTab)
    vals = Split(dat, vbTab)
    ReadTpl = True
End Function

Function DoOp(spec)
    Dim parts, tb, cs, tpl, g, pairs, p, nD, nI, nM, names, vals, k, v, overrides, eqp, fld, ord, c, line
    DoOp = False
    parts = Split(spec, "|", 5)
    tb = parts(1)
    cs = parts(2)
    tpl = parts(3)
    overrides = ";" & parts(4)
    log.WriteLine "---------- " & tb & " " & cs & " tpl=" & tpl
    Go "SE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = tb
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2000
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "10"
    session.findById("wnd[0]/usr/chkGD-EDIT").Selected = True
    Err.Clear
    On Error GoTo 0
    If Not SetSelByDesc("상담번호", cs) Then
        log.WriteLine "  FAIL could not set CSLTNO criterion"
        Exit Function
    End If
    Dim itp, itv
    itp = InStr(overrides, ";CSLTNO_ITEM=")
    If itp > 0 And (tb = "ZCRM_CICSVC" Or tb = "ZCRM_REQ") Then
        itv = Mid(overrides, itp + 13)
        If InStr(itv, ";") > 0 Then itv = Left(itv, InStr(itv, ";") - 1)
        If Not SetSelByDesc("상담번호의 seq(행)번호", itv) Then
            log.WriteLine "  FAIL could not set CSLTNO_ITEM criterion"
            Exit Function
        End If
        log.WriteLine "  item criterion=" & itv
    End If
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 3000
    Err.Clear
    On Error GoTo 0
    If CancelPopup("exec_" & cs) Then
        log.WriteLine "  FAIL popup after execute"
        Exit Function
    End If
    If Not Exists("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
        log.WriteLine "  FAIL no edit grid. screen=" & session.Info.ScreenNumber & " sbar=" & Sbar()
        Exit Function
    End If
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    log.WriteLine "  rows before=" & g.RowCount
    If g.RowCount <> 1 Then
        log.WriteLine "  FAIL expected exactly 1 row"
        Exit Function
    End If
    On Error Resume Next
    pairs = Split(parts(4), ";")
    For Each p In pairs
        eqp = InStr(p, "=")
        If eqp > 1 Then
            If Left(p, eqp - 1) <> "CSLTNO_ITEM" And Left(p, eqp - 1) <> "CSLTNO" Then
                g.modifyCell 0, Left(p, eqp - 1), Mid(p, eqp + 1)
                If Err.Number <> 0 Then
                    log.WriteLine "  ERR modifyCell " & p & " : " & Err.Description
                    Err.Clear
                End If
            End If
        End If
    Next
    g.pressEnter
    WScript.Sleep 1500
    Err.Clear
    On Error GoTo 0
    If CancelPopup("op_" & cs) Then
        log.WriteLine "  FAIL popup during input (not saved)"
        Exit Function
    End If
    log.WriteLine "  sbar after input: " & Sbar()
    If session.findById("wnd[0]/sbar").MessageType = "E" Then
        log.WriteLine "  FAIL error before save"
        Exit Function
    End If
    On Error Resume Next
    Set ord = g.ColumnOrder
    line = ""
    For c = 0 To ord.Count - 1
        line = line & ord(c) & "=" & g.GetCellValue(0, ord(c)) & " | "
        Err.Clear
    Next
    On Error GoTo 0
    log.WriteLine "  grid row0: " & line
    If Trim(g.GetCellValue(0, "CSLTNO")) <> cs Then
        log.WriteLine "  FAIL CSLTNO in grid differs"
        Exit Function
    End If
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S19_" & TAG & "_" & tb & "_" & cs & "_" & itv & "_before_save.png", 2
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    Err.Clear
    Dim w8
    For w8 = 1 To 20
        WScript.Sleep 500
        If Exists("wnd[1]") Then Exit For
    Next
    log.WriteLine "  after save: wait=" & (w8 * 500) & "ms screen=" & session.Info.ScreenNumber & " wnd1=" & Exists("wnd[1]") & " sbar=" & Sbar() & " rows=" & g.RowCount
    Err.Clear
    On Error GoTo 0
    nD = ""
    nI = ""
    nM = ""
    If Exists("wnd[1]/usr/txtGD_MOD_NR") Then
        nD = Trim(session.findById("wnd[1]/usr/txtGD_DEL_NR").Text)
        nI = Trim(session.findById("wnd[1]/usr/txtGD_INS_NR").Text)
        nM = Trim(session.findById("wnd[1]/usr/txtGD_MOD_NR").Text)
        session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\S19_" & TAG & "_" & tb & "_" & cs & "_" & itv & "_saved.png", 2
    End If
    log.WriteLine "  save result: deleted=" & nD & " inserted=" & nI & " modified=" & nM & " sbar=" & Sbar()
    CancelPopup "save_" & cs
    If nD <> "0" Or nI <> "0" Or nM <> "1" Then
        log.WriteLine "  FAIL counts differ (expected deleted=0 inserted=0 modified=1)"
        Exit Function
    End If
    log.WriteLine "  SAVED"
    DoOp = True
End Function

nFail = 0
For i = 1 To WScript.Arguments.Count - 1
    If Not DoOp(WScript.Arguments(i)) Then nFail = nFail + 1
    Go "SE16N"
    CancelPopup "leave_" & i
    EditOff
Next
Go "SE16N"
CancelPopup "final"
EditOff
On Error Resume Next
log.WriteLine "edit checkbox final=" & session.findById("wnd[0]/usr/chkGD-EDIT").Selected
Err.Clear
On Error GoTo 0
Go ""
log.WriteLine "failed=" & nFail
log.Close
session.findById("wnd[0]").close
WScript.Echo "S19 " & TAG & " failed=" & nFail
If nFail > 0 Then WScript.Quit 20
WScript.Quit 0
