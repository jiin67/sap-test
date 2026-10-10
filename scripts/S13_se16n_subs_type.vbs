' S13_se16n_subs_type.vbs : set ZCRM_CICSVC_NEW-SUBS_TYPE for given CSLTNO (item 1) via SE16N edit.
'   User-approved exception to CLAUDE.md section 12 (2026-09-11): SE16N edit ONLY for the listed CSLTNO.
'   args: <dry|apply> <newval: S | _ (= blank)> <csltno> [<csltno> ...]
'     dry   : opens the edit grid, verifies 1 row + expected old value, changes NOTHING.
'     apply : same checks, then modifies SUBS_TYPE of that single row and saves.
'   Expected old value: newval=S -> blank, newval=_ -> S. Any mismatch -> skip that CSLTNO (no save).
'   After save SE16N shows an info popup "변경 완료" with deleted/inserted/modified counts;
'   only 0/0/1 counts as success. Any other popup is logged and cancelled, never confirmed.
'   Edit checkbox is forced OFF at the end.
'   File must be saved as CP949 (Korean literals). Do not edit it with a UTF-8 editor.
'   log: logs\S13_<mode>_<yyyymmdd_hhnn>.txt (UTF-16)   exit: 0 all ok / 20 some CSLTNO failed
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Dim MODE, NEWVAL, OLDVAL, nFail, stamp

If WScript.Arguments.Count < 3 Then
    WScript.Echo "usage: S13_se16n_subs_type.vbs <dry|apply> <S|_> <csltno> ..."
    WScript.Quit 1
End If
MODE = LCase(WScript.Arguments(0))
If MODE <> "dry" And MODE <> "apply" Then WScript.Quit 1
If WScript.Arguments(1) = "S" Then
    NEWVAL = "S"
    OLDVAL = ""
ElseIf WScript.Arguments(1) = "_" Then
    NEWVAL = ""
    OLDVAL = "S"
Else
    WScript.Quit 1
End If

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
stamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2)
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S13_" & MODE & "_" & stamp & ".txt", True, True)
log.WriteLine "mode=" & MODE & " newval=[" & NEWVAL & "] expected old=[" & OLDVAL & "]"

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

' log + cancel any popup. Returns True if one was there.
Function CancelPopup(where)
    Dim w, c
    CancelPopup = False
    If Not Exists("wnd[1]") Then Exit Function
    CancelPopup = True
    Set w = session.findById("wnd[1]")
    log.WriteLine "  POPUP (" & where & ") title=[" & w.Text & "]"
    On Error Resume Next
    For Each c In w.findById("usr").Children
        log.WriteLine "    " & c.Type & " " & c.Name & " [" & c.Text & "]"
    Next
    w.HardCopy "C:\sap-test\results\shots\S13_" & stamp & "_popup_" & where & ".png", 2
    w.sendVKey 12
    WScript.Sleep 1000
    Err.Clear
    On Error GoTo 0
End Function

Sub EditOff()
    Dim ed
    On Error Resume Next
    Set ed = session.findById("wnd[0]/usr/chkGD-EDIT")
    If Err.Number = 0 Then ed.Selected = False
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

Function DoOne(csltno)
    Dim g, v, item, cs, ok, chg, nDel, nIns, nMod
    DoOne = False
    log.WriteLine "---------- CSLTNO " & csltno
    Go "SE16N"
    If session.Info.Program <> "SAPLSE16N" Then
        log.WriteLine "  FAIL not in SE16N: " & session.Info.Program
        Exit Function
    End If
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "ZCRM_CICSVC_NEW"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2000
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "10"
    session.findById("wnd[0]/usr/chkGD-EDIT").Selected = True
    If Err.Number <> 0 Then
        log.WriteLine "  FAIL setup " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If Not SetSelByDesc("상담번호", csltno) Then
        log.WriteLine "  FAIL could not set CSLTNO criterion"
        Exit Function
    End If
    If Not SetSelByDesc("상담번호의 seq(행)번호", "1") Then
        log.WriteLine "  FAIL could not set CSLTNO_ITEM criterion"
        Exit Function
    End If
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
    If CancelPopup("exec_" & csltno) Then
        log.WriteLine "  FAIL unexpected popup after execute"
        Exit Function
    End If
    log.WriteLine "  after exec screen=" & session.Info.ScreenNumber & " sbar=" & Sbar()
    If Not Exists("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
        log.WriteLine "  FAIL result grid not found"
        Exit Function
    End If
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    On Error Resume Next
    cs = g.GetCellValue(0, "CSLTNO")
    item = g.GetCellValue(0, "CSLTNO_ITEM")
    v = g.GetCellValue(0, "SUBS_TYPE")
    chg = g.GetCellChangeable(0, "SUBS_TYPE")
    If Err.Number <> 0 Then
        log.WriteLine "  (cell read note: " & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
    log.WriteLine "  rows=" & g.RowCount & " CSLTNO=[" & cs & "] ITEM=[" & item & "] SUBS_TYPE=[" & v & "] changeable=[" & chg & "]"
    ok = (g.RowCount = 1) And (Trim(cs) = csltno) And (Trim(item) = "1") And (Trim(v) = OLDVAL)
    If Not ok Then
        log.WriteLine "  FAIL precheck (need rows=1, same CSLTNO, item 1, SUBS_TYPE=[" & OLDVAL & "])"
        Exit Function
    End If
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S13_" & stamp & "_" & csltno & "_before.png", 2
    If MODE = "dry" Then
        log.WriteLine "  DRY OK - nothing changed"
        DoOne = True
        Exit Function
    End If

    ' ---- apply: change the single cell and save ----
    On Error Resume Next
    g.modifyCell 0, "SUBS_TYPE", NEWVAL
    g.currentCellRow = 0
    g.currentCellColumn = "SUBS_TYPE"
    g.pressEnter
    WScript.Sleep 800
    If Err.Number <> 0 Then
        log.WriteLine "  FAIL modifyCell " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    If CancelPopup("modify_" & csltno) Then
        log.WriteLine "  FAIL popup after modify (not saved)"
        Exit Function
    End If
    log.WriteLine "  cell now=[" & g.GetCellValue(0, "SUBS_TYPE") & "] sbar=" & Sbar()
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
    nDel = "" : nIns = "" : nMod = ""
    If Exists("wnd[1]/usr/txtGD_MOD_NR") Then
        nDel = Trim(session.findById("wnd[1]/usr/txtGD_DEL_NR").Text)
        nIns = Trim(session.findById("wnd[1]/usr/txtGD_INS_NR").Text)
        nMod = Trim(session.findById("wnd[1]/usr/txtGD_MOD_NR").Text)
        log.WriteLine "  save result popup: deleted=" & nDel & " inserted=" & nIns & " modified=" & nMod
    End If
    If CancelPopup("save_" & csltno) Then
        If Not (nDel = "0" And nIns = "0" And nMod = "1") Then
            log.WriteLine "  FAIL unexpected popup/counts after save - check DB"
            Exit Function
        End If
    End If
    log.WriteLine "  after save screen=" & session.Info.ScreenNumber & " sbar=" & Sbar()
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S13_" & stamp & "_" & csltno & "_after.png", 2
    log.WriteLine "  SAVED"
    DoOne = True
End Function

nFail = 0
For i = 2 To WScript.Arguments.Count - 1
    If Not DoOne(WScript.Arguments(i)) Then nFail = nFail + 1
    ' leave the edit list; a pending-changes popup here would mean an unsaved edit
    Go "SE16N"
    If CancelPopup("leave_" & i) Then
        log.WriteLine "  WARN popup when leaving list"
        Go "SE16N"
    End If
    EditOff
Next

' make sure edit mode is OFF on the SE16N initial screen, then leave
Go "SE16N"
CancelPopup "final"
EditOff
log.WriteLine "edit checkbox final=" & session.findById("wnd[0]/usr/chkGD-EDIT").Selected
Go ""
log.WriteLine "failed=" & nFail
log.Close
session.findById("wnd[0]").close
WScript.Echo "S13 " & MODE & " failed=" & nFail
If nFail > 0 Then WScript.Quit 20
WScript.Quit 0
