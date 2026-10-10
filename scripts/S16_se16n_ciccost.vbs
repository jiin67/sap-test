' S16_se16n_ciccost.vbs : insert / delete ZCRM_CICCOST_M test rows through SE16N edit mode.
'   User-approved (2026-09-13): 4 temporary rows so that 기술료/출장비 aggregation can be verified.
'   Only the 4 test CSLTNO are allowed. Each operation opens SE16N with an EXACT 상담번호 criterion
'   (the table is large - never load it unrestricted), changes one row and saves.
'   args: <tag> <op> [<op> ...]
'     ins:<CSLTNO>:<FIELD>=<value>,<FIELD>=<value>,...   (row must not exist yet)
'     del:<CSLTNO>                                        (row must exist, exactly 1)
'   Save result popup must show inserted=1 (ins) or deleted=1 (del); otherwise FAIL.
'   Popups other than the save result are logged and cancelled. Edit checkbox forced OFF at the end.
'   File must be saved as CP949 (Korean literal "상담번호").
'   log: logs\S16_<tag>.txt (UTF-16)   exit: 0 ok / 20 some op failed
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Dim TAG, nFail, allowed, today

allowed = "|2026063003075|2026063003277|2026070100389|2026070101688|"
If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: S16_se16n_ciccost.vbs <tag> ins:CSLTNO:F=v,F=v | del:CSLTNO ..."
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
For i = 1 To WScript.Arguments.Count - 1
    If InStr(allowed, "|" & Split(WScript.Arguments(i), ":")(1) & "|") = 0 Then
        WScript.Echo "CSLTNO not allowed: " & WScript.Arguments(i)
        WScript.Quit 1
    End If
Next
today = Year(Now) & "." & Right("0" & Month(Now), 2) & "." & Right("0" & Day(Now), 2)

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
' pick the KDQ connection explicitly (KDP/KDD may be open in the same SAP GUI)
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
Set log = fso.CreateTextFile("C:\sap-test\logs\S16_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG

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
    w.HardCopy "C:\sap-test\results\shots\S16_" & TAG & "_popup_" & where & ".png", 2
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
    Dim parts, kind, cs, g, pairs, p, kv, nD, nI, nM, expD, expI
    DoOp = False
    parts = Split(spec, ":", 3)
    kind = parts(0)
    cs = parts(1)
    log.WriteLine "---------- " & spec
    Go "SE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "ZCRM_CICCOST_M"
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
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 3000
    Err.Clear
    On Error GoTo 0
    If CancelPopup("exec") Then
        log.WriteLine "  FAIL popup after execute"
        Exit Function
    End If
    If Not Exists("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
        log.WriteLine "  FAIL no edit grid. screen=" & session.Info.ScreenNumber & " sbar=" & Sbar()
        Exit Function
    End If
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    log.WriteLine "  rows before=" & g.RowCount

    If kind = "ins" Then
        If g.RowCount <> 0 Then
            log.WriteLine "  FAIL row already exists"
            Exit Function
        End If
        On Error Resume Next
        g.pressToolbarButton "&LOCAL&APPEND"
        WScript.Sleep 800
        g.modifyCell 0, "CSLTNO", cs
        pairs = Split(parts(2), ",")
        For Each p In pairs
            kv = Split(p, "=")
            g.modifyCell 0, kv(0), kv(1)
            If Err.Number <> 0 Then
                log.WriteLine "  ERR modifyCell " & kv(0) & " : " & Err.Description
                Err.Clear
            End If
        Next
        g.modifyCell 0, "ERNAM", "ZTEST0922"
        g.modifyCell 0, "ERDAT", today
        g.pressEnter
        WScript.Sleep 1000
        Err.Clear
        On Error GoTo 0
        expD = "0" : expI = "1"
    ElseIf kind = "del" Then
        If g.RowCount <> 1 Then
            log.WriteLine "  FAIL expected exactly 1 row"
            Exit Function
        End If
        If Trim(g.GetCellValue(0, "ERNAM")) <> "ZTEST0922" Then
            log.WriteLine "  FAIL row was not created by this test (ERNAM=" & g.GetCellValue(0, "ERNAM") & ")"
            Exit Function
        End If
        On Error Resume Next
        g.selectedRows = "0"
        g.pressToolbarButton "&LOCAL&DELETE_ROW"
        WScript.Sleep 800
        Err.Clear
        On Error GoTo 0
        expD = "1" : expI = "0"
    Else
        log.WriteLine "  FAIL unknown op"
        Exit Function
    End If
    If CancelPopup("op") Then
        log.WriteLine "  FAIL popup during op (not saved)"
        Exit Function
    End If
    If session.findById("wnd[0]/sbar").MessageType = "E" Then
        log.WriteLine "  FAIL error before save: " & Sbar()
        Exit Function
    End If
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S16_" & TAG & "_" & cs & "_before_save.png", 2
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
    If nD <> expD Or nI <> expI Then
        log.WriteLine "  FAIL counts differ (expected deleted=" & expD & " inserted=" & expI & ")"
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
WScript.Echo "S16 " & TAG & " failed=" & nFail
If nFail > 0 Then WScript.Quit 20
WScript.Quit 0
