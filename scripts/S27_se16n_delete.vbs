' S27_se16n_delete.vbs : DELETE incomplete TEST rows (own ZTEST data only) through SE16N edit mode. Based on S19.
'   2026-10-02: the first mk1002_e2e.ps1 run had a broken template path, so ZCRM_REQ / ZCRM_CICCOST_M / _SUB / _ADD /
'   ZCRM_CHRG_CNFRM rows of 2026100290001~010 were saved with empty template columns. They are removed and re-inserted.
'   args: <tag> del|<TABLE>|<CSLTNO>|<expected ERNAM or ->  ...
'     - CSLTNO must start with 202610029 (this test set only), table must be in the allowed list
'     - opens SE16N edit mode with EXACT CSLTNO; 1..3 rows allowed; every row's ERNAM must equal the expected value
'       (skip the check with -); all rows are selected, deleted with the grid's delete-row button and saved
'     - save popup must show deleted=<rows> inserted=0 modified=0
'   log: logs\S27_<tag>.txt (UTF-16). exit 0 ok / 1 bad args / 20 some op failed. Pure ASCII + ChrW().
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, pp
Dim TAG, nFail, allowedTb, KO_CSLTNO

KO_CSLTNO = ChrW(49345) & ChrW(45812) & ChrW(48264) & ChrW(54840)   ' sang-dam-beon-ho
allowedTb = "|ZCRM_CICCOST_M|ZCRM_CICCOST_SUB|ZCRM_CICCOST_ADD|ZCRM_REQ|ZCRM_CHRG_CNFRM|ZCRM_CICSVC|ZCRM_CICSVC_NEW|"
If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: S27_se16n_delete.vbs <tag> del|TABLE|CSLTNO|ERNAM ..."
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
For i = 1 To WScript.Arguments.Count - 1
    pp = Split(WScript.Arguments(i), "|")
    If UBound(pp) < 3 Then
        WScript.Echo "bad op: " & WScript.Arguments(i)
        WScript.Quit 1
    End If
    If pp(0) <> "del" Or InStr(allowedTb, "|" & pp(1) & "|") = 0 Or Left(pp(2), 9) <> "202610029" Or (Len(pp(2)) <> 13 And Len(pp(2)) <> 12) Then
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
Set log = fso.CreateTextFile("C:\sap-test\logs\S27_" & TAG & ".txt", True, True)
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
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
End Sub

Function CancelPopup(where)
    Dim w
    CancelPopup = False
    If Not Exists("wnd[1]") Then Exit Function
    CancelPopup = True
    Set w = session.findById("wnd[1]")
    log.WriteLine "  POPUP (" & where & ") [" & w.Text & "]"
    On Error Resume Next
    w.HardCopy "C:\sap-test\results\shots\S27_" & TAG & "_popup_" & where & ".png", 2
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

Function DoOp(spec)
    Dim parts, tb, cs, ern, g, nD, nI, nM, r, nr, v, bid, k, delId, w8
    DoOp = False
    parts = Split(spec, "|")
    tb = parts(1)
    cs = parts(2)
    ern = parts(3)
    log.WriteLine "---------- DEL " & tb & " " & cs & " expect ERNAM=" & ern
    Go "SE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = tb
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2000
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "10"
    session.findById("wnd[0]/usr/chkGD-EDIT").Selected = True
    Err.Clear
    On Error GoTo 0
    If Not SetSelByDesc(KO_CSLTNO, cs) Then
        log.WriteLine "  FAIL could not set CSLTNO criterion"
        Exit Function
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
    nr = g.RowCount
    log.WriteLine "  rows=" & nr
    If nr < 1 Or nr > 3 Then
        log.WriteLine "  FAIL unexpected row count"
        Exit Function
    End If
    On Error Resume Next
    For r = 0 To nr - 1
        v = Trim(g.GetCellValue(r, "CSLTNO"))
        If v <> cs Then
            log.WriteLine "  FAIL row " & r & " CSLTNO=" & v
            On Error GoTo 0
            Exit Function
        End If
        If ern <> "-" Then
            v = Trim(g.GetCellValue(r, "ERNAM"))
            If Err.Number <> 0 Or v <> ern Then
                log.WriteLine "  FAIL row " & r & " ERNAM=" & v & " err=" & Err.Number
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
        End If
    Next
    delId = ""
    For k = 0 To g.ToolbarButtonCount - 1
        bid = g.GetToolbarButtonId(k)
        If InStr(UCase(bid), "DELETE") > 0 Then delId = bid
    Next
    Err.Clear
    If delId = "" Then
        log.WriteLine "  FAIL no delete button on grid toolbar"
        On Error GoTo 0
        Exit Function
    End If
    log.WriteLine "  delete button=" & delId
    g.selectedRows = "0-" & (nr - 1)
    WScript.Sleep 300
    g.pressToolbarButton delId
    WScript.Sleep 1500
    Err.Clear
    On Error GoTo 0
    If CancelPopup("del_" & cs) Then
        log.WriteLine "  FAIL popup after delete (not saved)"
        Exit Function
    End If
    log.WriteLine "  rows after delete (unsaved)=" & g.RowCount & " sbar=" & Sbar()
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S27_" & TAG & "_" & tb & "_" & cs & "_before_save.png", 2
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    Err.Clear
    For w8 = 1 To 20
        WScript.Sleep 500
        If Exists("wnd[1]") Then Exit For
    Next
    Err.Clear
    On Error GoTo 0
    nD = "" : nI = "" : nM = ""
    If Exists("wnd[1]/usr/txtGD_MOD_NR") Then
        nD = Trim(session.findById("wnd[1]/usr/txtGD_DEL_NR").Text)
        nI = Trim(session.findById("wnd[1]/usr/txtGD_INS_NR").Text)
        nM = Trim(session.findById("wnd[1]/usr/txtGD_MOD_NR").Text)
    End If
    log.WriteLine "  save result: deleted=" & nD & " inserted=" & nI & " modified=" & nM & " sbar=" & Sbar()
    CancelPopup "save_" & cs
    If nD <> CStr(nr) Or nI <> "0" Or nM <> "0" Then
        log.WriteLine "  FAIL counts differ (expected deleted=" & nr & ")"
        Exit Function
    End If
    log.WriteLine "  DELETED " & nr
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
Go ""
log.WriteLine "failed=" & nFail
log.Close
session.findById("wnd[0]").close
WScript.Echo "S27 " & TAG & " failed=" & nFail
If nFail > 0 Then WScript.Quit 20
WScript.Quit 0
