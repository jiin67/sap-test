' S21_se16n_setfield.vbs : change ONE field of ONE test row through SE16N edit mode (pure ASCII file, Korean via ChrW).
'   User-approved 2026-09-18 (ZCRMR171 scenario): SUBS_TYPE re-validation tests need ZCRM_CICSVC-SUBS_TYPE flipped and restored.
'   args: <tag> <table> <csltno> <item> <field> <old|_> <new|_>      ( _ = blank )
'   allowed: table ZCRM_CICSVC, field SUBS_TYPE, CSLTNO 20260918900xx or 2026091490002/3
'   The row is selected by exact CSLTNO + CSLTNO_ITEM, must be exactly 1 row and hold <old>. Save popup must show modified=1.
'   log: logs\S21_<tag>.txt (UTF-16)  shots: results\shots\S21_<tag>_before.png / _after.png   exit 0 ok / 20 failed / 99 guard
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Dim TAG, TBL, CS, ITEM, FLD, OLDV, NEWV, D_CS, D_ITEM, g, cur, nD, nI, nM, rcode

If WScript.Arguments.Count < 7 Then
    WScript.Echo "usage: S21 <tag> <table> <csltno> <item> <field> <old|_> <new|_>"
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
TBL = UCase(WScript.Arguments(1))
CS = WScript.Arguments(2)
ITEM = WScript.Arguments(3)
FLD = UCase(WScript.Arguments(4))
OLDV = WScript.Arguments(5)
NEWV = WScript.Arguments(6)
If OLDV = "_" Then OLDV = ""
If NEWV = "_" Then NEWV = ""
If TBL <> "ZCRM_CICSVC" Or (FLD <> "SUBS_TYPE" And FLD <> "SUBSWRTMONTH") Then WScript.Quit 1
If (Left(CS, 9) <> "202609189" And Left(CS, 9) <> "202609199" And Left(CS, 9) <> "202609219") And CS <> "2026091490002" And CS <> "2026091490003" Then WScript.Quit 1
D_CS = ChrW(49345) & ChrW(45812) & ChrW(48264) & ChrW(54840)
D_ITEM = D_CS & ChrW(51032) & " seq(" & ChrW(54665) & ")" & ChrW(48264) & ChrW(54840)

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
Set log = fso.CreateTextFile("C:\sap-test\logs\S21_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG & " " & TBL & " " & CS & "/" & ITEM & " " & FLD & " [" & OLDV & "] -> [" & NEWV & "]"
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

Sub Go(tc)
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tc
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2000
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

Sub Finish(code)
    On Error Resume Next
    If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
    Go "SE16N"
    If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
    session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    log.WriteLine "edit checkbox final=" & session.findById("wnd[0]/usr/chkGD-EDIT").Selected
    Go ""
    Err.Clear
    On Error GoTo 0
    log.WriteLine "exit code=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Echo "S21 " & TAG & " exit=" & code
    WScript.Quit code
End Sub

Go "SE16N"
On Error Resume Next
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = TBL
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "10"
session.findById("wnd[0]/usr/chkGD-EDIT").Selected = True
Err.Clear
On Error GoTo 0
If Not SetSelByDesc(D_CS, CS) Then
    log.WriteLine "FAIL criterion CSLTNO"
    Finish 20
End If
If Not SetSelByDesc(D_ITEM, ITEM) Then
    log.WriteLine "FAIL criterion CSLTNO_ITEM"
    Finish 20
End If
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
If Exists("wnd[1]") Or Not Exists("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
    log.WriteLine "FAIL no edit grid / popup"
    Finish 20
End If
Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
If g.RowCount <> 1 Then
    log.WriteLine "FAIL rows=" & g.RowCount
    Finish 20
End If
cur = Trim(g.GetCellValue(0, FLD))
log.WriteLine "current [" & cur & "] csltno=" & g.GetCellValue(0, "CSLTNO") & " item=" & g.GetCellValue(0, "CSLTNO_ITEM")
If Trim(g.GetCellValue(0, "CSLTNO")) <> CS Or cur <> OLDV Then
    log.WriteLine "FAIL row/old value mismatch"
    Finish 20
End If
On Error Resume Next
g.setCurrentCell 0, FLD
WScript.Sleep 500
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S21_" & TAG & "_before.png", 2
g.modifyCell 0, FLD, NEWV
g.pressEnter
WScript.Sleep 1200
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S21_" & TAG & "_after.png", 2
session.findById("wnd[0]/tbar[0]/btn[11]").press
Err.Clear
For i = 1 To 20
    WScript.Sleep 500
    If Exists("wnd[1]") Then Exit For
Next
On Error GoTo 0
nD = "" : nI = "" : nM = ""
If Exists("wnd[1]/usr/txtGD_MOD_NR") Then
    nD = Trim(session.findById("wnd[1]/usr/txtGD_DEL_NR").Text)
    nI = Trim(session.findById("wnd[1]/usr/txtGD_INS_NR").Text)
    nM = Trim(session.findById("wnd[1]/usr/txtGD_MOD_NR").Text)
End If
log.WriteLine "save result: deleted=" & nD & " inserted=" & nI & " modified=" & nM
rcode = 0
If nD <> "0" Or nI <> "0" Or nM <> "1" Then
    log.WriteLine "FAIL counts"
    rcode = 20
Else
    log.WriteLine "SAVED"
End If
Finish rcode
