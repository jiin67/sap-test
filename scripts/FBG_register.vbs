' FBG_register.vbs : KDD ONLY. Register ZCRMF03_RFC_ZCRM_CHRG_CNFRM in FBGENDAT (mode B) and set user parameter FBGENDAT=X in SU3, with screenshots.
Option Explicit
Const LOGNAME = "FBG_register"
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
Const FM = "ZCRMF03_RFC_ZCRM_CHRG_CNFRM"
Dim g, r, c, cols, line

' ================= Step 1-3 : FBGENDAT registration =================
Tcode "/nSE38"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "FBGENDAT"
Shot "FBG_01_se38"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 2500
Status "FBGENDAT sel:"
If session.Info.Program <> "FBGENDAT" Then Finish 11
session.findById("wnd[0]/usr/txtPA_FNAME").Text = FM
session.findById("wnd[0]/usr/txtPA_MODE").Text = "B"
session.findById("wnd[0]/usr/radPA_INS").Select
WScript.Sleep 500
log.WriteLine "check FNAME=[" & T("wnd[0]/usr/txtPA_FNAME") & "] MODE=[" & T("wnd[0]/usr/txtPA_MODE") & "] INS=" & session.findById("wnd[0]/usr/radPA_INS").Selected
If T("wnd[0]/usr/txtPA_FNAME") <> FM Or T("wnd[0]/usr/txtPA_MODE") <> "B" Or Not session.findById("wnd[0]/usr/radPA_INS").Selected Then Finish 12
Shot "FBG_02_sel_filled"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 2500
Status "after INS F8:"
If Exists("wnd[1]") Then
    log.WriteLine "POPUP [" & T("wnd[1]") & "]"
    DumpTree session.findById("wnd[1]"), 4, 0
    session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\FBG_03_popup.png", 2
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 1500
    Status "after popup:"
End If
DumpTree session.findById("wnd[0]/usr"), 4, 0
Shot "FBG_03_after_ins"

' ---- Step 4 : verify with 'display active entries'
Tcode "/nSE38"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "FBGENDAT"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 2000
session.findById("wnd[0]/usr/txtPA_FNAME").Text = FM
session.findById("wnd[0]/usr/radPA_READ").Select
Shot "FBG_04_sel_read"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 2500
Status "READ:"
Dim okReg : okReg = False
On Error Resume Next
Set g = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
If Err.Number = 0 Then
    Set cols = g.ColumnOrder
    log.WriteLine "  rows=" & g.RowCount
    For r = 0 To g.RowCount - 1
        line = ""
        For c = 0 To cols.Count - 1 : line = line & g.GetCellValue(r, cols(c)) & vbTab : Next
        log.WriteLine "  " & r & ": " & line
        If g.GetCellValue(r, "FNAME") = FM And g.GetCellValue(r, "MODE") = "B" Then okReg = True
    Next
End If
Err.Clear
On Error GoTo 0
log.WriteLine "REGISTERED=" & okReg
Shot "FBG_05_read_result"
If Not okReg Then Finish 13

' ================= Step 5 : SU3 user parameter FBGENDAT = X =================
Const GRID = "wnd[0]/usr/tabsTABSTRIP1/tabpPARAM/ssubMAINAREA:SAPLSUID_MAINTENANCE:1104/cntlG_PARAMETER_CONTAINER/shellcont/shell"
Tcode "/nSU3"
Status "SU3:"
If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : DumpTree session.findById("wnd[1]"), 3, 0 : session.findById("wnd[1]").sendVKey 0 : WScript.Sleep 800
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpPARAM").Select
WScript.Sleep 1500
Set g = session.findById(GRID)
Set cols = g.ColumnOrder
line = ""
For c = 0 To cols.Count - 1 : line = line & cols(c) & vbTab : Next
log.WriteLine "  param grid rows=" & g.RowCount & " cols=" & line
Dim blank, already : blank = -1 : already = -1
For r = 0 To g.RowCount - 1
    If r Mod 20 = 0 Then g.FirstVisibleRow = r
    If Trim(g.GetCellValue(r, "PARID")) = "FBGENDAT" Then already = r : log.WriteLine "  FBGENDAT already at row " & r & " value=[" & g.GetCellValue(r, "PARVA") & "]"
    If Trim(g.GetCellValue(r, "PARID")) = "" And blank < 0 Then blank = r
Next
log.WriteLine "  first blank row=" & blank & " already=" & already
g.FirstVisibleRow = 0
Shot "FBG_06_su3_param_before"
If already >= 0 Then
    log.WriteLine "  parameter exists - no change"
Else
    If blank < 0 Then
        log.WriteLine "  no blank row - try append button"
        DumpTree session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpPARAM"), 6, 0
        Finish 14
    End If
    g.FirstVisibleRow = blank
    g.ModifyCell blank, "PARID", "FBGENDAT"
    g.ModifyCell blank, "PARVA", "X"
    g.SetCurrentCell blank, "PARVA"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1500
    Status "after enter:"
    If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : DumpTree session.findById("wnd[1]"), 3, 0 : session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\FBG_07_su3_popup.png", 2 : session.findById("wnd[1]").sendVKey 0 : WScript.Sleep 800
    Set g = session.findById(GRID)
    ' re-locate the row (grid may re-sort)
    already = -1
    For r = 0 To g.RowCount - 1
        If r Mod 20 = 0 Then g.FirstVisibleRow = r
        If Trim(g.GetCellValue(r, "PARID")) = "FBGENDAT" Then already = r
    Next
    log.WriteLine "  after enter FBGENDAT row=" & already
    If already < 0 Then Finish 15
    g.FirstVisibleRow = already
    log.WriteLine "  PARID=[" & g.GetCellValue(already, "PARID") & "] PARVA=[" & g.GetCellValue(already, "PARVA") & "] text=[" & g.GetCellValue(already, cols(cols.Count - 1)) & "]"
    Shot "FBG_07_su3_param_entered"
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    WScript.Sleep 2000
    Status "after save:"
    If Exists("wnd[1]") Then log.WriteLine "POPUP [" & T("wnd[1]") & "]" : DumpTree session.findById("wnd[1]"), 3, 0 : session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\FBG_08_su3_save_popup.png", 2 : session.findById("wnd[1]").sendVKey 0 : WScript.Sleep 800
    Shot "FBG_08_su3_saved"
End If

' ---- verify via USR05 (display only)
Tcode "/nSE16N"
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "USR05"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "200"
If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500
If InStr(T("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0,2]"), "ID") > 0 Then
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").Text = "FBGENDAT"
    session.findById("wnd[0]").sendVKey 8
    WScript.Sleep 3000
    Status "USR05:"
    On Error Resume Next
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    For r = 0 To g.RowCount - 1
        log.WriteLine "  " & g.GetCellValue(r, "BNAME") & " " & g.GetCellValue(r, "PARID") & " [" & g.GetCellValue(r, "PARVA") & "]"
    Next
    Err.Clear
    On Error GoTo 0
    Shot "FBG_09_usr05_after"
End If
Finish 0
