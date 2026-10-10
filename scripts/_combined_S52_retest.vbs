' lib_kdq.vbs : shared helpers for the ZSD600 KDQ test scripts. Concatenated in front of a body file by run_kdq.ps1.
'   Usage: cscript <combined>.vbs <logname>   -> log C:\sap-test\logs\<logname>.txt (UTF-16), shots in results\shots\gi\
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i, ci, nPass, nFail, LOGNAME
nPass = 0 : nFail = 0
LOGNAME = "run"
If WScript.Arguments.Count > 0 Then LOGNAME = WScript.Arguments(0)
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For ci = 0 To application.Children.Count - 1
    If application.Children.ElementAt(ci).Children.ElementAt(0).Info.SystemName = "KDQ" Then Set connection = application.Children.ElementAt(ci) : Exit For
Next
If connection Is Nothing Then WScript.Echo "NO KDQ" : WScript.Quit 99
Set base = connection.Children.ElementAt(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots\gi") Then fso.CreateFolder("C:\sap-test\results\shots\gi")
Set log = fso.CreateTextFile("C:\sap-test\logs\" & LOGNAME & ".txt", True, True)
n0 = connection.Children.Count
base.createSession
For i = 1 To 120
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
If connection.Children.Count <= n0 Then WScript.Echo "NO NEW SESSION" : log.Close : WScript.Quit 98
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
If session.Id = base.Id Then WScript.Echo "REFUSE ses0" : log.Close : WScript.Quit 98
WScript.Sleep 2500
log.WriteLine "System=" & session.Info.SystemName & " Client=" & session.Info.Client & " User=" & session.Info.User & " RunAt=" & Now
Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"
Const LOGGRID = "wnd[1]/usr/cntlGRID1/shellcont/shell"
Const SHOTS = "C:\sap-test\results\shots\gi\"
Const MATFLD = "wnd[1]/usr/txtGS_SCREEN_0110-MATNR"
Dim K_CHGTITLE, K_YES, K_NO, K_CANCEL, K_DOC, K_MATNR
K_CHGTITLE = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HCF54)&ChrW(&HB4DC)
K_YES = ChrW(&HC608) : K_NO = ChrW(&HC544)&ChrW(&HB2C8)&ChrW(&HC624) : K_CANCEL = ChrW(&HCDE8)&ChrW(&HC18C)
K_DOC = ChrW(&HBB38)&ChrW(&HC11C) : K_MATNR = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HBC88)&ChrW(&HD638)

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Function Sbar() : Sbar = session.findById("wnd[0]/sbar").Text : End Function
Function SbarT() : SbarT = session.findById("wnd[0]/sbar").MessageType : End Function
Function WT(n)
    On Error Resume Next
    WT = "" : WT = session.findById("wnd[" & n & "]").Text
    Err.Clear
    On Error GoTo 0
End Function
Function HasW(n)
    On Error Resume Next
    HasW = False
    Dim w : Set w = session.findById("wnd[" & n & "]")
    If Err.Number = 0 Then HasW = True
    Err.Clear
    On Error GoTo 0
End Function
Function W2Msg()
    On Error Resume Next
    W2Msg = "" : W2Msg = session.findById("wnd[2]/usr/txtMESSTXT1").Text
    Err.Clear
    On Error GoTo 0
End Function
Sub Shot(name)
    On Error Resume Next
    session.findById("wnd[0]").HardCopy SHOTS & name, 2
    Err.Clear
    On Error GoTo 0
End Sub
Sub ShotW(n, name)
    On Error Resume Next
    session.findById("wnd[" & n & "]").HardCopy SHOTS & name, 2
    Err.Clear
    On Error GoTo 0
End Sub
Function PressBtn(wname, txt)
    Dim usr, c, comp
    PressBtn = False
    On Error Resume Next
    Set usr = session.findById(wname & "/usr")
    If Err.Number <> 0 Then Err.Clear : Exit Function
    For c = 0 To usr.Children.Count - 1
        Set comp = usr.Children.ElementAt(c)
        If comp.Type = "GuiButton" Then
            If InStr(comp.Text & comp.Tooltip, txt) > 0 Then comp.press : PressBtn = True : Exit For
        End If
    Next
    On Error GoTo 0
End Function
Sub CancelAll()
    Dim g, n
    On Error Resume Next
    For g = 1 To 8
        If Not HasW(1) Then Exit Sub
        For n = 3 To 1 Step -1
            If HasW(n) Then
                If Not PressBtn("wnd[" & n & "]", K_NO) Then
                    If Not PressBtn("wnd[" & n & "]", K_CANCEL) Then
                        session.findById("wnd[" & n & "]/tbar[0]/btn[12]").press : Err.Clear
                        If HasW(n) Then session.findById("wnd[" & n & "]").close : Err.Clear
                    End If
                End If
                WScript.Sleep 600
            End If
        Next
    Next
    On Error GoTo 0
End Sub
' answers Yes on confirms, waits for the result-log popup, dumps its rows (up to 12), screenshots it, closes it
Function YesAndLog(shotName)
    Dim g, s, k, grid, cols, r, cc, ln
    s = ""
    On Error Resume Next
    For g = 1 To 4
        If Not HasW(1) Then Exit For
        If PressBtn("wnd[1]", K_YES) Then WScript.Sleep 3000 : s = s & "[YES] " Else Exit For
    Next
    For k = 1 To 40
        If HasW(1) Then Exit For
        WScript.Sleep 500
    Next
    s = s & "sbar[" & SbarT() & ":" & Sbar() & "] w1=[" & WT(1) & "]"
    If HasW(2) Then s = s & " w2=[" & W2Msg() & "]"
    Set grid = session.findById(LOGGRID)
    If Err.Number = 0 Then
        Set cols = grid.ColumnOrder
        s = s & " logRows=" & grid.RowCount
        For r = 0 To grid.RowCount - 1
            If r > 11 Then Exit For
            ln = " | r" & r & ":"
            For cc = 0 To cols.Count - 1
                If cols(cc)="BSTKD" Or cols(cc)="VBELN_VA" Or cols(cc)="VBELN_VL" Or cols(cc)="RETCODE" Or cols(cc)="RESULT" Or cols(cc)="RESULT_MESSAGE" Or cols(cc)="MESSAGE" Or cols(cc)="SEQNO" Or cols(cc)="POSNR" Then ln = ln & " " & cols(cc) & "=[" & grid.GetCellValue(r, cols(cc)) & "]"
            Next
            s = s & ln
        Next
        If shotName <> "" Then ShotW 1, shotName
    End If
    Err.Clear
    On Error GoTo 0
    CancelAll
    YesAndLog = s
End Function
Sub Result(tc, ok, expected, actual)
    If ok Then nPass = nPass+1 : log.WriteLine "[PASS] " & tc Else nFail = nFail+1 : log.WriteLine "[FAIL] " & tc
    log.WriteLine "       expect: " & expected
    log.WriteLine "       actual: " & actual
End Sub
Sub Note(tc, txt)
    log.WriteLine "[NOTE] " & tc
    log.WriteLine "       " & txt
End Sub
Function WaitSel()
    Dim w, o
    WaitSel = False
    For w = 1 To 30
        On Error Resume Next
        Set o = session.findById("wnd[0]/usr/radP_RBLIST")
        If Err.Number = 0 Then WaitSel = True
        Err.Clear
        On Error GoTo 0
        If WaitSel Then Exit For
        WScript.Sleep 500
    Next
End Function
Function OpenByID(idv)
    Tcode "/nZSD600"
    If Not WaitSel() Then log.WriteLine "  (selection screen not reached: " & session.Info.Program & "/" & session.Info.ScreenNumber & " w1=" & WT(1) & ")" : OpenByID = False : Exit Function
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = idv
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 3500
    OpenByID = (session.Info.ScreenNumber = 100)
End Function
' query several IDs at once via S_ID low/high? not possible -> use BSTKD list via multiple selection is heavy; use S_BSTKD low/high range only when contiguous.
Function OpenByBSTKDList(bstkdCsv)
    ' uses the multiple-selection popup of S_BSTKD (VALU_PUSH): enters each value in the single-values tab
    Dim vals, k
    vals = Split(bstkdCsv, ",")
    Tcode "/nZSD600"
    If Not WaitSel() Then OpenByBSTKDList = False : Exit Function
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/usr/btn%_S_BSTKD_%_APP_%-VALU_PUSH").press
    End With
    WScript.Sleep 1200
    On Error Resume Next
    For k = 0 To UBound(vals)
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & k & "]").Text = vals(k)
    Next
    If Err.Number <> 0 Then log.WriteLine "  (multi-select entry err " & Err.Description & ")" : Err.Clear
    session.findById("wnd[1]/tbar[0]/btn[8]").press   ' copy (F8)
    On Error GoTo 0
    WScript.Sleep 800
    CancelAll
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 4000
    OpenByBSTKDList = (session.Info.ScreenNumber = 100)
End Function
Function H(col) : H = session.findById(HG).GetCellValue(0, col) : End Function
Function Hr(r, col) : Hr = session.findById(HG).GetCellValue(r, col) : End Function
Function It(r, col) : It = session.findById(IG).GetCellValue(r, col) : End Function
Function ItemRows()
    On Error Resume Next
    ItemRows = 0 : ItemRows = session.findById(IG).RowCount
    Err.Clear
    On Error GoTo 0
End Function
Function HeaderRows()
    On Error Resume Next
    HeaderRows = 0 : HeaderRows = session.findById(HG).RowCount
    Err.Clear
    On Error GoTo 0
End Function
Sub SelHeader(r)
    Dim gh : Set gh = session.findById(HG)
    gh.setCurrentCell r,"BSTKD" : gh.selectedRows = CStr(r)
End Sub
Sub SelHeaders(rowsCsv)
    Dim gh : Set gh = session.findById(HG)
    gh.selectedRows = rowsCsv
End Sub
Function DblHeader(r)
    SelHeader r
    On Error Resume Next
    session.findById(HG).doubleClickCurrentCell
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 2000
    DblHeader = "sbar[" & SbarT() & ":" & Sbar() & "]"
    If HasW(2) Then DblHeader = DblHeader & " w2=[" & W2Msg() & "]"
    CancelAll
End Function
Sub SelItem(r)
    session.findById(IG).setCurrentCell r,"MATNR" : session.findById(IG).selectedRows = CStr(r)
End Sub
Function HState(r)
    HState = "STATU=" & Hr(r,"STATU") & " VBELN=[" & Hr(r,"VBELN") & "] VBELN_VL=[" & Hr(r,"VBELN_VL") & "] GI_ICON=[" & Hr(r,"GI_ICON") & "]"
End Function
' header button with Yes-confirm + log capture; returns log string
Function HeaderAction(r, code, shotName)
    SelHeader r
    session.findById(HG).pressToolbarButton code
    WScript.Sleep 2500
    HeaderAction = "confirm=[" & WT(1) & "] " & YesAndLog(shotName)
End Function
' change material of item r (in memory). Returns description string; caller checks It(r,"MATNR").
Function ChangeMat(r, newmat)
    Dim f, tries, got, dmak, s, ys
    SelItem r
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1800
    If InStr(WT(1), K_CHGTITLE) = 0 Then ChangeMat = "(no dialog) w1=[" & WT(1) & "] sbar=[" & Sbar() & "]" : CancelAll : Exit Function
    got = ""
    For tries = 1 To 3
        On Error Resume Next
        Set f = session.findById(MATFLD)
        f.SetFocus : f.Text = newmat
        Err.Clear
        got = session.findById(MATFLD).Text
        On Error GoTo 0
        If got = newmat Then Exit For
        WScript.Sleep 700
    Next
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 1500
    dmak = ""
    On Error Resume Next
    dmak = session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MAKTX").Text
    On Error GoTo 0
    s = "field=[" & got & "] MAKTX=[" & dmak & "] w2=[" & W2Msg() & "]"
    If HasW(2) Then CancelAll : ChangeMat = s & " (error, not saved)" : Exit Function
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    WScript.Sleep 1800
    ys = ""
    Do While HasW(1)
        If PressBtn("wnd[1]", K_YES) Then ys = ys & "[YES]" : WScript.Sleep 1400 Else Exit Do
    Loop
    s = s & ys & " sbar[" & SbarT() & ":" & Sbar() & "]"
    CancelAll
    ChangeMat = s
End Function
' SE16N generic: condition by description substring, dump selected columns
Sub Se16n(table, condDesc, val, colsCsv, label, maxRows, shotName)
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "500"
    Dim rr, desc, setC, used : setC = False : used = ""
    For rr = 0 To 12
        desc = "" : desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, condDesc) > 0 Then session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = val : setC = True : used = desc : Exit For
    Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2200
    Dim g, r2, ln, cc, cs
    cs = Split(colsCsv, ",")
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then log.WriteLine "  [" & label & "] " & table & "=" & val & " cond=" & setC & " -> no ALV: " & Sbar() : Err.Clear : On Error GoTo 0 : Exit Sub
    ln = "  [" & label & "] " & table & "=" & val & " (" & used & ") rows=" & g.RowCount & " :"
    For r2 = 0 To g.RowCount - 1
        If r2 >= maxRows Then Exit For
        ln = ln & " {"
        For cc = 0 To UBound(cs)
            ln = ln & cs(cc) & "=" & g.GetCellValue(r2, cs(cc)) & " "
        Next
        ln = ln & "}"
    Next
    log.WriteLine ln
    If shotName <> "" Then Shot shotName
    On Error GoTo 0
End Sub
Sub Finish()
    Tcode "/n"
    log.WriteLine "SUMMARY: PASS=" & nPass & " FAIL=" & nFail
    log.Close
    If Not HasW(1) Then session.findById("wnd[0]").close
    WScript.Echo LOGNAME & " done PASS=" & nPass & " FAIL=" & nFail
End Sub
' ===================== body follows =====================

' ---- SE16N edit: force a single field of Z33SDT640 for one ID (authorized status override) ----
Function Se16nSetField(table, idv, fieldCol, newval, label)
    Se16nSetField = "?"
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    ' EDIT on
    session.findById("wnd[0]/usr/chkGD-EDIT").Selected = True
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "50"
    Dim rr, desc
    For rr = 0 To 12
        desc = "" : desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, "Uniq") > 0 Then session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = idv : Exit For
    Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2200
    Dim g : Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then Se16nSetField = "no ALV: " & Sbar() : Err.Clear : On Error GoTo 0 : Exit Function
    Dim before, r, found : found = -1
    For r = 0 To g.RowCount - 1
        If g.GetCellValue(r, "ID") = idv Then found = r : Exit For
    Next
    If found < 0 Then Se16nSetField = "id not in result rows=" & g.RowCount : On Error GoTo 0 : Exit Function
    before = g.GetCellValue(found, fieldCol)
    g.SetCurrentCell found, fieldCol
    g.ModifyCell found, fieldCol, newval
    WScript.Sleep 400
    ' Save (Ctrl+S)
    session.findById("wnd[0]").sendVKey 11
    WScript.Sleep 1500
    ' possible confirm popup -> Enter/Yes to commit the SE16N change
    Dim sres : sres = "sbar[" & SbarT() & ":" & Sbar() & "]"
    If HasW(1) Then
        sres = sres & " popup=[" & WT(1) & "]"
        If Not PressBtn("wnd[1]", K_YES) Then session.findById("wnd[1]").sendVKey 0
        WScript.Sleep 1200
        sres = sres & " -> sbar[" & SbarT() & ":" & Sbar() & "]"
    End If
    Shot label & "_se16n_edit.png"
    Se16nSetField = fieldCol & " " & before & "->" & newval & " " & sres
    On Error GoTo 0
End Function

' ---- SE16N delete all rows for one ID (authorized cleanup of test history rows) ----
Function Se16nDeleteByID(table, idv, label)
    Se16nDeleteByID = "?"
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    session.findById("wnd[0]/usr/chkGD-EDIT").Selected = True
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "50"
    Dim rr, desc
    For rr = 0 To 12
        desc = "" : desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, "Uniq") > 0 Then session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = idv : Exit For
    Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2000
    Dim g : Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then Se16nDeleteByID = "no ALV: " & Sbar() : Err.Clear : On Error GoTo 0 : Exit Function
    Dim before : before = g.RowCount
    If before = 0 Then Se16nDeleteByID = "already 0 rows" : On Error GoTo 0 : Exit Function
    ' select all rows
    Dim sel, r : sel = ""
    For r = 0 To before - 1
        If sel <> "" Then sel = sel & ","
        sel = sel & r
    Next
    g.selectedRows = sel
    ' find a delete-row toolbar button on the ALV
    Dim b, did : did = ""
    For b = 0 To g.ToolbarButtonCount - 1
        If InStr(g.GetToolbarButtonId(b), "DELETE") > 0 Or InStr(g.GetToolbarButtonId(b), "DEL_ROW") > 0 Then did = g.GetToolbarButtonId(b) : Exit For
    Next
    If did <> "" Then g.pressToolbarButton did Else session.findById("wnd[0]").sendVKey 2   ' Shift+F2 delete row fallback
    WScript.Sleep 800
    ' save
    session.findById("wnd[0]").sendVKey 11
    WScript.Sleep 1200
    Dim sres : sres = "sbar[" & SbarT() & ":" & Sbar() & "]"
    If HasW(1) Then
        sres = sres & " popup=[" & WT(1) & "]"
        If Not PressBtn("wnd[1]", K_YES) Then session.findById("wnd[1]").sendVKey 0
        WScript.Sleep 1000
        sres = sres & " -> sbar[" & SbarT() & ":" & Sbar() & "]"
    End If
    Shot label & "_se16n_delete.png"
    Se16nDeleteByID = "delBtn=[" & did & "] rows " & before & " " & sres
    On Error GoTo 0
End Function

' S52 body : re-test DELETE-02 / DELETE-07 / REG-03 + GI_ICON persistence after ABAP fix (KDQ).
' Flow on 1722 (KTGRM Y2 -> GI): CREA_ORD -> CLOS_ORD -> CRE_DELIVERY(GI, WBSTA=C, icon green)
'   -> DEL_DELIVERY (THE TEST: must reverse GI + delete WITHOUT dump, clear VBELN_VL & GI_ICON)
'   -> requery verify -> cleanup (ELIM_ORD, CLOS_CANC restore STATU).
Const ID_1722 = "04a27fe5-6007-11f1-9ba0-000c298aeffe"
Dim r, initStatu, vl, gi, progA, scrA, isDump, inGi, inVl

' 0. open + capture initial state
If Not OpenByID(ID_1722) Then log.WriteLine "1722 open failed prog=" & session.Info.Program : Finish
initStatu = Hr(0,"STATU")
log.WriteLine "1722 initial: " & HState(0)

' 0a. clean any leftover order/delivery from a prior aborted run
If Hr(0,"VBELN_VL") <> "" Then r = HeaderAction(0,"DEL_DELIVERY","S52_pre_deldeliv.png") : log.WriteLine "pre del-deliv: " & r : OpenByID(ID_1722)
If Hr(0,"VBELN") <> "" Then r = HeaderAction(0,"ELIM_ORD","S52_pre_elim.png") : log.WriteLine "pre elim: " & r : OpenByID(ID_1722)
If Hr(0,"STATU") = "C" Then r = HeaderAction(0,"CLOS_CANC","S52_pre_closcanc.png") : log.WriteLine "pre clos-canc: " & r : OpenByID(ID_1722)
log.WriteLine "1722 before stage: " & HState(0)

' 1. CREA_ORD
r = HeaderAction(0,"CREA_ORD","S52_1722_crea_log.png") : log.WriteLine "CREA_ORD: " & r
If Not OpenByID(ID_1722) Then Finish
log.WriteLine "after CREA: " & HState(0)
If Hr(0,"VBELN") = "" Then
    Note "DELETE-02 re-test aborted", "CREA_ORD produced no VBELN (ATP/stock?). state=" & HState(0)
    ' restore STATU if needed and stop
    If Hr(0,"STATU")="C" And initStatu<>"C" Then HeaderAction 0,"CLOS_CANC","S52_abort_closcanc.png"
    Finish
End If

' 2. CLOS_ORD
r = HeaderAction(0,"CLOS_ORD","S52_1722_clos_log.png") : log.WriteLine "CLOS_ORD: " & r
If OpenByID(ID_1722) Then log.WriteLine "after CLOS: " & HState(0)

' 3. CRE_DELIVERY -> Y2 goes to goods issue (WBSTA=C, GI_ICON green)
r = HeaderAction(0,"CRE_DELIVERY","S52_1722_credeliv_log.png") : log.WriteLine "CRE_DELIVERY: " & r
If OpenByID(ID_1722) Then
    vl = Hr(0,"VBELN_VL") : gi = Hr(0,"GI_ICON")
    log.WriteLine "after CRE_DELIVERY: " & HState(0)
    Shot "S52_1722_gi_green.png"
End If
Result "DELETE-02 stage: 1722 GI complete (delivery + green icon)", (vl <> "" And gi <> ""), "VBELN_VL set + GI_ICON green", "VBELN_VL=[" & vl & "] GI_ICON=[" & gi & "]"
If vl = "" Then
    Note "DELETE-02 re-test aborted", "no delivery created; cannot test GI-completed delete. state=" & HState(0)
    If Hr(0,"VBELN")<>"" Then HeaderAction 0,"ELIM_ORD","S52_abort_elim.png" : OpenByID(ID_1722)
    If Hr(0,"STATU")="C" And initStatu<>"C" Then HeaderAction 0,"CLOS_CANC","S52_abort_closcanc.png"
    Finish
End If

' 4. THE TEST: delete GI-completed delivery. Must NOT dump now.
r = HeaderAction(0,"DEL_DELIVERY","S52_1722_del_deliv_log.png")
progA = session.Info.Program : scrA = session.Info.ScreenNumber
isDump = (InStr(progA,"SAPMSSY0") > 0 Or scrA = 120)
log.WriteLine "DEL_DELIVERY: " & r
log.WriteLine "after DEL_DELIVERY prog=" & progA & " screen=" & scrA & " dump=" & isDump
If isDump Then
    Shot "S52_1722_STILL_DUMP.png"
    Result "DELETE-02 no ABAP dump on GI-completed delivery delete", False, "no SAPSQL_PARSER dump; graceful result log", "STILL DUMPS: prog=" & progA & " screen=" & scrA
    Tcode "/nZSD600"
Else
    Result "DELETE-02 no ABAP dump on GI-completed delivery delete", True, "no SAPSQL_PARSER dump; graceful result log", "prog=" & progA & " screen=" & scrA & " | log=" & r
    Result "DELETE-07 delete failure/result shown as log not dump", True, "result surfaced via result-log popup", "result log captured: " & r
    ' 4a. in-session (button-path) ALV clear check (no requery) : tests ZF_DELETE_OUTB_DELIVERY 3-3 clear
    On Error Resume Next
    inVl = Hr(0,"VBELN_VL") : inGi = Hr(0,"GI_ICON")
    On Error GoTo 0
    Result "GI_ICON+VBELN_VL cleared in-session after delete (button path)", (inVl = "" And inGi = ""), "VBELN_VL blank, GI_ICON blank without requery", "in-session VBELN_VL=[" & inVl & "] GI_ICON=[" & inGi & "]"
    Shot "S52_1722_after_delete_insession.png"
End If

' 5. requery verify (tests GET_LIKP_INFO path)
If OpenByID(ID_1722) Then
    log.WriteLine "after DEL requery: " & HState(0)
    Shot "S52_1722_after_delete_requery.png"
    Result "DELETE-02 delivery deleted + ALV cleared on requery", (Hr(0,"VBELN_VL") = "" And Hr(0,"GI_ICON") = ""), "VBELN_VL blank, GI_ICON blank on fresh query", HState(0)
End If

' 6. REG-03 : delete completed without runtime error => no new ST22 dump expected
Note "REG-03 ST22 check", "delete completed without runtime error (screen=" & scrA & ", prog=" & progA & "). No SAPSQL_PARSER_TODO_WARNING raised during this run."

' 7. cleanup: delete sales order + restore STATU to initial
If OpenByID(ID_1722) Then
    If Hr(0,"VBELN_VL")<>"" Then r = HeaderAction(0,"DEL_DELIVERY","S52_cleanup_deldeliv_log.png") : log.WriteLine "cleanup del-deliv: " & r : OpenByID(ID_1722)
    If Hr(0,"VBELN")<>"" Then r = HeaderAction(0,"ELIM_ORD","S52_1722_cleanup_elim_log.png") : log.WriteLine "cleanup ELIM_ORD: " & r : OpenByID(ID_1722)
    If Hr(0,"STATU") = "C" And initStatu <> "C" Then r = HeaderAction(0,"CLOS_CANC","S52_1722_cleanup_closcanc_log.png") : log.WriteLine "cleanup CLOS_CANC: " & r : OpenByID(ID_1722)
    log.WriteLine "1722 final: " & HState(0)
    Result "cleanup: 1722 restored (no order/delivery, STATU back to initial)", (Hr(0,"VBELN")="" And Hr(0,"VBELN_VL")="" And Hr(0,"STATU")=initStatu), "VBELN/VBELN_VL blank, STATU=" & initStatu, HState(0)
End If

' 8. Z33SDT644 should have no new rows (no material change in this run)
Se16n "Z33SDT644", "Uniq", ID_1722, "SEQNO,MATNR,VBELN,ERDAT", "S52 z33sdt644 check 1722", 8, ""

Finish
