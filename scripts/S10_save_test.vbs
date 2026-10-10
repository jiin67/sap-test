' S10_save_test.vbs
'   TC-07 / TC-08 / TC-09 : enter values on ONE designated row and press save.
'   The row is addressed by CSLTNO + MATNR_PRT + TAGNO1, never by row index.
'   Records the before state, the inputs, the status-bar message and the after
'   state, and captures screenshots at each step.
'
'   *** THIS SCRIPT CHANGES THE DATABASE for TC07 / TC08. ***
'   TC09* cases are designed to be rejected by the program's own validation.
'
'   Leaving the screen always answers the save-confirmation popup with the
'   SECOND button (No) so nothing is written by accident.
'
'   args: <case> <OLD|NEW> <CSLTNO> <MATNR_PRT> <TAGNO1>
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, f, n0, i
Dim SHOTDIR, RAWDIR, STAMP, RSNCODE
SHOTDIR = "C:\sap-test\results\shots\"
RAWDIR = "C:\sap-test\results\raw\"
STAMP = "ZTEST_20260910"
RSNCODE = "Z02"

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

Set connection = Nothing
For i = 0 To application.Children.Count - 1
    Set base = application.Children.ElementAt(i).Children.ElementAt(0)
    If base.Info.SystemName = "KDQ" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then
    WScript.Echo "GUARD: no KDQ connection"
    WScript.Quit 99
End If
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then
    WScript.Echo "GUARD: " & base.Info.SystemName & " - save tests run on KDQ only"
    WScript.Quit 99
End If

Dim CASEID, PROGSEL, CSLTNO, MATNR, TAGNO, ARGS
Set ARGS = WScript.Arguments
If ARGS.Count < 5 Then
    WScript.Echo "usage: S10_save_test.vbs <case> <OLD|NEW> <CSLTNO> <MATNR_PRT> <TAGNO1>"
    WScript.Quit 1
End If
CASEID = UCase(ARGS(0))
PROGSEL = UCase(ARGS(1))
CSLTNO = ARGS(2)
MATNR = ARGS(3)
TAGNO = ARGS(4)

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder "C:\sap-test\results\shots"
Set log = fso.CreateTextFile("C:\sap-test\logs\S10_" & CASEID & "_" & PROGSEL & ".txt", True, True)
Set f = fso.CreateTextFile(RAWDIR & "save_" & CASEID & "_" & PROGSEL & ".tsv", True, True)
log.WriteLine "case=" & CASEID & " prog=" & PROGSEL & " CSLTNO=" & CSLTNO & " MATNR=" & MATNR & " TAGNO=" & TAGNO
f.WriteLine "#CASE" & vbTab & CASEID
f.WriteLine "#PROG" & vbTab & PROGSEL
f.WriteLine "#KEY" & vbTab & CSLTNO & vbTab & MATNR & vbTab & TAGNO

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
If connection.Children.Count <= n0 Then
    log.WriteLine "FAIL cannot create session"
    log.Close
    WScript.Quit 2
End If
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
log.WriteLine "work session=" & session.Id

Function NewTcode()
    Dim p, ts, ln, v
    v = ""
    p = "C:\sap-test\scripts\config_local.txt"
    On Error Resume Next
    If fso.FileExists(p) Then
        Set ts = fso.OpenTextFile(p, 1)
        Do While Not ts.AtEndOfStream
            ln = Trim(ts.ReadLine)
            If InStr(ln, "TCODE_NEW=") = 1 Then v = Trim(Mid(ln, 11))
        Loop
        ts.Close
    End If
    Err.Clear
    On Error GoTo 0
    NewTcode = v
End Function

Sub Shot(nm)
    On Error Resume Next
    session.findById("wnd[0]").HardCopy SHOTDIR & nm & ".png", 2
    Err.Clear
    On Error GoTo 0
End Sub

Sub ShotPopup(nm)
    Dim w
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    w.HardCopy SHOTDIR & nm & ".png", 2
    Err.Clear
    On Error GoTo 0
End Sub

' Log every button of the current popup, then press the requested one.
' whichButton = 1 (yes) / 2 (no) / 3 (cancel). 0 = only log, press nothing.
Function PopupButtons(tag, whichButton)
    Dim w, j, b, txts
    txts = ""
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        PopupButtons = ""
        Exit Function
    End If
    log.WriteLine tag & " POPUP [" & w.Text & "]"
    f.WriteLine "#POPUP_" & tag & vbTab & w.Text
    For j = 1 To 3
        Set b = session.findById("wnd[1]/usr/btnBUTTON_" & j)
        If Err.Number = 0 Then txts = txts & j & "=[" & b.Text & "] "
        Err.Clear
    Next
    log.WriteLine tag & " buttons " & txts
    f.WriteLine "#POPUPBTN_" & tag & vbTab & txts
    If whichButton > 0 Then
        Set b = session.findById("wnd[1]/usr/btnBUTTON_" & whichButton)
        If Err.Number = 0 Then
            log.WriteLine tag & " pressing BUTTON_" & whichButton & " [" & b.Text & "]"
            b.press
            WScript.Sleep 1500
        Else
            Err.Clear
            log.WriteLine tag & " no BUTTON_" & whichButton & ", sending F12"
            Set w = session.findById("wnd[1]")
            If Err.Number = 0 Then
                w.sendVKey 12
                WScript.Sleep 1200
            End If
            Err.Clear
        End If
    End If
    Err.Clear
    On Error GoTo 0
    PopupButtons = txts
End Function

' leave the screen discarding any unsaved change
Sub LeaveDiscard()
    Dim w, guard
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1800
    Err.Clear
    On Error GoTo 0
    guard = 0
    Do
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        On Error GoTo 0
        ShotPopup "S10_" & CASEID & "_" & PROGSEL & "_9_exitpopup" & guard
        PopupButtons "EXIT", 2
        guard = guard + 1
    Loop While guard < 3
End Sub

Function SetField(nm, val)
    Dim o
    On Error Resume Next
    Set o = session.findById("wnd[0]/usr/ctxt" & nm)
    If Err.Number <> 0 Then
        Err.Clear
        Set o = session.findById("wnd[0]/usr/txt" & nm)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        SetField = False
        Exit Function
    End If
    o.Text = val
    SetField = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Dim tc, ok
On Error Resume Next
If PROGSEL = "NEW" Then
    tc = NewTcode()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tc
Else
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
End If
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2400
session.findById("wnd[0]/usr/radPA_R01").Select
session.findById("wnd[0]/usr/radPA_NCH").Select
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
Err.Clear
On Error GoTo 0
ok = SetField("SO_RCDT-LOW", "")
ok = SetField("SO_RCDT-HIGH", "")
ok = SetField("SO_PRDRE-LOW", "2026.07.01")
ok = SetField("SO_PRDRE-HIGH", "2026.07.15")
ok = SetField("SO_CSLTN-LOW", CSLTNO)
ok = SetField("SO_CSLTN-HIGH", "")
Shot "S10_" & CASEID & "_" & PROGSEL & "_1_selection"

On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2500
Err.Clear
On Error GoTo 0

Dim g
On Error Resume Next
Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
If Err.Number <> 0 Then
    Err.Clear
    log.WriteLine "FAIL no grid  sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    f.WriteLine "#RESULT" & vbTab & "NO_GRID"
    f.Close
    log.Close
    WScript.Quit 10
End If
On Error GoTo 0
log.WriteLine "rows=" & g.RowCount

Dim r, target
target = -1
On Error Resume Next
For r = 0 To g.RowCount - 1
    If g.GetCellValue(r, "MATNR_PRT") = MATNR And g.GetCellValue(r, "TAGNO1") = TAGNO Then
        target = r
        Exit For
    End If
    Err.Clear
Next
Err.Clear
On Error GoTo 0
If target < 0 Then
    log.WriteLine "FAIL target row not found"
    f.WriteLine "#RESULT" & vbTab & "ROW_NOT_FOUND"
    LeaveDiscard
    f.Close
    log.Close
    WScript.Quit 11
End If
log.WriteLine "target row=" & target
f.WriteLine "#TARGETROW" & vbTab & target

Dim SNAPCOLS
SNAPCOLS = Array("CSLTNO", "REQNO", "RECPTNO", "RECPTDATE", "SRV_QTY", "CF_QTY", "NCF_QTY", _
                 "NCFT", "NCF_RSN_VT", "RSPC", "RSPCT", "WEARDATS", "REMARK", "REMARK1", _
                 "REMARK2", "MATNR_PRT", "TAGNO1", "KUNNR_SVC", "VKBUR", "CONFDAT", "PRICETYPET")

Sub Snap(tag)
    Dim k, v
    On Error Resume Next
    For k = 0 To UBound(SNAPCOLS)
        v = g.GetCellValue(target, SNAPCOLS(k))
        If Err.Number <> 0 Then
            v = "<n/a>"
            Err.Clear
        End If
        f.WriteLine "#" & tag & vbTab & SNAPCOLS(k) & vbTab & v
    Next
    v = session.findById("wnd[0]/usr/txtGS_HEADER-REMARK1").Text
    If Err.Number = 0 Then f.WriteLine "#" & tag & vbTab & "HDR_REMARK1" & vbTab & v
    Err.Clear
    v = session.findById("wnd[0]/usr/txtGS_HEADER-REMARK2").Text
    If Err.Number = 0 Then f.WriteLine "#" & tag & vbTab & "HDR_REMARK2" & vbTab & v
    Err.Clear
    On Error GoTo 0
End Sub

Snap "BEFORE"
Shot "S10_" & CASEID & "_" & PROGSEL & "_2_before"

Dim inCF, inNCF, inRSN, inRSPC, inWEAR, inHDR1, inHDR2, srvqty
On Error Resume Next
srvqty = g.GetCellValue(target, "SRV_QTY")
Err.Clear
On Error GoTo 0
If Trim(srvqty) = "" Then srvqty = "1"

inCF = ""
inNCF = ""
inRSN = ""
inRSPC = ""
inWEAR = ""
inHDR1 = STAMP
inHDR2 = STAMP

Select Case CASEID
    Case "TC07"
        inCF = "1"
        inRSPC = "1100"
        inWEAR = "2026.07.15"
    Case "TC08"
        inNCF = "1"
        inRSN = RSNCODE
        inWEAR = "2026.07.15"
    Case "TC09-01"
        inCF = CStr(CInt(srvqty) + 5)
        inRSPC = "1100"
        inWEAR = "2026.07.15"
    Case "TC09-02"
        inNCF = "1"
        inWEAR = "2026.07.15"
    Case "TC09-03"
        inRSN = RSNCODE
        inWEAR = "2026.07.15"
    Case "TC09-04"
        inCF = "1"
        inWEAR = "2026.07.15"
    Case "TC09-05"
        inCF = "1"
        inRSPC = "1100"
        inWEAR = "2026.07.15"
        inHDR1 = ""
        inHDR2 = ""
    Case "TC09-06"
        inCF = "1"
        inRSPC = "1100"
        inWEAR = ""
    Case Else
        log.WriteLine "FAIL unknown case " & CASEID
        LeaveDiscard
        f.Close
        log.Close
        WScript.Quit 12
End Select

f.WriteLine "#INPUT" & vbTab & "CF_QTY" & vbTab & inCF
f.WriteLine "#INPUT" & vbTab & "NCF_QTY" & vbTab & inNCF
f.WriteLine "#INPUT" & vbTab & "NCF_RSN_VT" & vbTab & inRSN
f.WriteLine "#INPUT" & vbTab & "RSPC" & vbTab & inRSPC
f.WriteLine "#INPUT" & vbTab & "WEARDATS" & vbTab & inWEAR
f.WriteLine "#INPUT" & vbTab & "HDR_REMARK1" & vbTab & inHDR1
f.WriteLine "#INPUT" & vbTab & "HDR_REMARK2" & vbTab & inHDR2
log.WriteLine "inputs CF=" & inCF & " NCF=" & inNCF & " RSN=" & inRSN & _
              " RSPC=" & inRSPC & " WEAR=" & inWEAR & " HDR=" & inHDR1

Sub Put(col, val)
    Dim rb
    If val = "" Then Exit Sub
    On Error Resume Next
    g.setCurrentCell target, col
    g.modifyCell target, col, val
    If Err.Number <> 0 Then
        log.WriteLine "  ERR modifyCell " & col & "=" & val & " : " & Err.Description
        f.WriteLine "#PUTERR" & vbTab & col & vbTab & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    rb = g.GetCellValue(target, col)
    If Err.Number <> 0 Then
        rb = "<readerr>"
        Err.Clear
    End If
    log.WriteLine "  put " & col & "=" & val & " readback=[" & rb & "]"
    f.WriteLine "#PUTBACK" & vbTab & col & vbTab & val & vbTab & rb
    On Error GoTo 0
End Sub

Put "CF_QTY", inCF
Put "NCF_QTY", inNCF
Put "NCF_RSN_VT", inRSN
Put "RSPC", inRSPC
Put "WEARDATS", inWEAR

' Header fields are ordinary screen fields; the grid keeps its edits in the
' front end. Pressing Enter here would make the ALV redraw from the backend
' and throw the typed values away, so we go straight to save.
On Error Resume Next
session.findById("wnd[0]/usr/txtGS_HEADER-REMARK1").Text = inHDR1
session.findById("wnd[0]/usr/txtGS_HEADER-REMARK2").Text = inHDR2
Err.Clear
On Error GoTo 0
Shot "S10_" & CASEID & "_" & PROGSEL & "_3_entered"
On Error Resume Next
f.WriteLine "#SBAR_AFTER_ENTER" & vbTab & session.findById("wnd[0]/sbar").Text
Err.Clear
On Error GoTo 0
Snap "ENTERED"

log.WriteLine "pressing save"
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 3500
If Err.Number <> 0 Then
    log.WriteLine "  ERR save " & Err.Number & " " & Err.Description
    Err.Clear
End If
On Error GoTo 0

Dim sbarTxt, sbarType, btnTxts
On Error Resume Next
sbarTxt = session.findById("wnd[0]/sbar").Text
sbarType = session.findById("wnd[0]/sbar").MessageType
Err.Clear
On Error GoTo 0
f.WriteLine "#SBAR_AFTER_SAVE" & vbTab & sbarTxt
f.WriteLine "#SBAR_TYPE" & vbTab & sbarType
log.WriteLine "after save sbar=[" & sbarTxt & "] type=" & sbarType
Shot "S10_" & CASEID & "_" & PROGSEL & "_4_after_save"
ShotPopup "S10_" & CASEID & "_" & PROGSEL & "_4_popup"
btnTxts = PopupButtons("AFTERSAVE", 0)
If btnTxts <> "" Then
    ' this is the "저장 하시겠습니까?" confirmation - answer YES so the
    ' program's own validation runs and we can observe its message
    btnTxts = PopupButtons("SAVECONFIRM", 1)
    WScript.Sleep 2500
    On Error Resume Next
    f.WriteLine "#SBAR_AFTER_CONFIRM" & vbTab & session.findById("wnd[0]/sbar").Text
    f.WriteLine "#SBARTYPE_AFTER_CONFIRM" & vbTab & session.findById("wnd[0]/sbar").MessageType
    log.WriteLine "after confirm sbar=[" & session.findById("wnd[0]/sbar").Text & "] type=" & session.findById("wnd[0]/sbar").MessageType
    Err.Clear
    On Error GoTo 0
    Shot "S10_" & CASEID & "_" & PROGSEL & "_4b_after_confirm"
    ShotPopup "S10_" & CASEID & "_" & PROGSEL & "_4c_confirm_popup"
    ' any further popup is a message - read it, then close with Enter
    On Error Resume Next
    Dim w4
    Set w4 = session.findById("wnd[1]")
    If Err.Number = 0 Then
        f.WriteLine "#POPUP_AFTERCONFIRM" & vbTab & w4.Text
        log.WriteLine "popup after confirm [" & w4.Text & "]"
        w4.sendVKey 0
        WScript.Sleep 1500
    End If
    Err.Clear
    On Error GoTo 0
Else
    On Error Resume Next
    Dim w3
    Set w3 = session.findById("wnd[1]")
    If Err.Number = 0 Then
        w3.sendVKey 0
        WScript.Sleep 1200
    End If
    Err.Clear
    On Error GoTo 0
End If

On Error Resume Next
Set g = session.findById("wnd[0]/usr/cntlCON1/shellcont/shell")
If Err.Number <> 0 Then
    Err.Clear
    f.WriteLine "#AFTER" & vbTab & "GRID_GONE"
Else
    On Error GoTo 0
    Snap "AFTER"
End If
Err.Clear
On Error GoTo 0
Shot "S10_" & CASEID & "_" & PROGSEL & "_5_after"

LeaveDiscard

On Error Resume Next
log.WriteLine "final prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber
Err.Clear
On Error GoTo 0
f.Close
log.WriteLine "done"
log.Close
On Error Resume Next
session.findById("wnd[0]").close
Err.Clear
On Error GoTo 0
WScript.Echo "S10 " & CASEID & " " & PROGSEL & " done"
