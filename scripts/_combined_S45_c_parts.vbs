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

' S45 body : parts/boiler delivery+GI flow.
'   1722 (parts, KTGRM Y2): CLOS_CANC -> CREA_ORD -> CLOS_ORD -> CRE_DELIVERY (expect delivery + GI: DIV-02, ICON-03, ALV-04/05) -> VBUP
'   2392 (boiler Y1):        CLOS_CANC -> CREA_ORD -> CLOS_ORD -> (delivery in mixed run)
'   2827 (B, LGORT blank):   CREA_ORD -> CLOS_ORD -> (delivery in mixed run, expect error)
'   DEL-03: 2766 (no SO) -> CRE_DELIVERY blocked
'   DEL-06/REG-04: 2392 + 2827 + 1757(dup) selected together -> CRE_DELIVERY
'   ALV-06: re-query 1722 after GI ; ALV-08 layout popups
Const ID_1722 = "04a27fe5-6007-11f1-9ba0-000c298aeffe"
Const ID_2392 = "04a5f20c-8a85-11f1-9bc0-000c298aeffe"
Const ID_2827 = "04bf872f-a20b-11f1-9bce-000c298aeffe"
Const ID_2766 = "04ad1bf8-a070-11f1-9bcd-000c298aeffe"
Dim r, msg, vb, vl

Function PrepSO(idv, tag)
    ' returns VBELN after CLOS_CANC(if C) -> CREA_ORD -> CLOS_ORD
    PrepSO = ""
    If Not OpenByID(idv) Then log.WriteLine "  " & tag & " not found" : Exit Function
    log.WriteLine tag & " start: " & HState(0)
    If Hr(0,"STATU") = "C" And Hr(0,"VBELN") = "" Then
        r = HeaderAction(0, "CLOS_CANC", tag & "_clos_canc_log.png")
        If Not OpenByID(idv) Then Exit Function
        log.WriteLine "  after CLOS_CANC: " & HState(0)
    End If
    If Hr(0,"VBELN") = "" Then
        msg = DblHeader(0)
        log.WriteLine "  dblclick ATP: " & msg
        Shot tag & "_atp_dblclick.png"
        r = HeaderAction(0, "CREA_ORD", tag & "_crea_ord_log.png")
        log.WriteLine "  CREA_ORD: " & r
        If Not OpenByID(idv) Then Exit Function
    End If
    vb = Hr(0,"VBELN")
    log.WriteLine "  VBELN=" & vb & " " & HState(0)
    If vb <> "" And Hr(0,"STATU") <> "C" Then
        r = HeaderAction(0, "CLOS_ORD", tag & "_clos_ord_log.png")
        If Not OpenByID(idv) Then Exit Function
        log.WriteLine "  after CLOS_ORD: " & HState(0)
    End If
    PrepSO = Hr(0,"VBELN")
End Function

' ---------- 1722 ----------
vb = PrepSO(ID_1722, "C2_1722")
Result "PRE 1722 sales order + close", (vb <> ""), "VBELN, STATU C", "VBELN=" & vb & " " & HState(0)
If vb <> "" Then
    r = HeaderAction(0, "CRE_DELIVERY", "DIV_02_1722_cre_delivery_log.png")
    Shot "DIV_02_1722_after_button_inplace.png"
    Dim giInPlace : giInPlace = Hr(0,"GI_ICON")
    Dim vlInPlace : vlInPlace = Hr(0,"VBELN_VL")
    vl = ""
    If OpenByID(ID_1722) Then vl = Hr(0,"VBELN_VL") : r = r & " requery: " & HState(0)
    Shot "DIV_02_1722_requery.png"
    Result "DIV-02/DEL-02 1722 (parts KTGRM Y2) delivery + goods issue", (vl <> "" And Hr(0,"GI_ICON") <> ""), "delivery created AND GI posted (GI_ICON green)", r
    Result "ALV-05 1722 GI_ICON updated in place right after action", (giInPlace <> "" And vlInPlace <> ""), "GI_ICON green without re-query", "inplace VBELN_VL=[" & vlInPlace & "] GI_ICON=[" & giInPlace & "]"
    Result "ALV-06/ICON-04 1722 GI_ICON after fresh re-query", (Hr(0,"GI_ICON") <> ""), "green icon from VBUP-WBSTA", "GI_ICON=[" & Hr(0,"GI_ICON") & "]"
    If vl <> "" Then Se16n "VBUP", K_DOC, vl, "VBELN,POSNR,WBSTA,LFSTA,FKSTA", "ICON-03 VBUP " & vl, 12, "ICON_03_1722_vbup.png"
End If

' ---------- 2392 ----------
vb = PrepSO(ID_2392, "C2_2392")
Result "PRE 2392 sales order + close", (vb <> ""), "VBELN, STATU C", "VBELN=" & vb & " " & HState(0)

' ---------- 2827 (B, LGORT blank) ----------
vb = PrepSO(ID_2827, "C2_2827")
Note "PRE 2827 (LGORT blank) sales order + close", "VBELN=" & vb & " " & HState(0)

' ---------- DEL-03 : 2766 no sales order ----------
If OpenByID(ID_2766) Then
    log.WriteLine "2766: " & HState(0)
    r = HeaderAction(0, "CRE_DELIVERY", "DEL_03_2766_no_so_log.png")
    vl = "?"
    If OpenByID(ID_2766) Then vl = Hr(0,"VBELN_VL")
    Result "DEL-03 delivery on order without sales order (2766) blocked", (vl = ""), "blocked, error log", r
End If

' ---------- DEL-06 / REG-04 : mixed 2392 + 2827 + 1757(dup) ----------
If OpenByBSTKDList("2392,2827,1757") Then
    Dim n, rows, k
    n = HeaderRows()
    log.WriteLine "mixed query rows=" & n
    rows = ""
    For k = 0 To n - 1
        log.WriteLine "  H" & k & ": BSTKD=" & Hr(k,"BSTKD") & " " & HState(k)
        If Hr(k,"VBELN") <> "" Then
            If rows <> "" Then rows = rows & ","
            rows = rows & k
        End If
    Next
    SelHeaders rows
    session.findById(HG).pressToolbarButton "CRE_DELIVERY"
    WScript.Sleep 3000
    r = "selected=[" & rows & "] confirm=[" & WT(1) & "] " & YesAndLog("DEL_06_mixed_log.png")
    Shot "DEL_06_mixed_after.png"
    log.WriteLine "  DEL-06: " & r
    Dim s2392, s2827
    If OpenByID(ID_2392) Then s2392 = HState(0) : Shot "DEL_06_2392_after.png"
    If OpenByID(ID_2827) Then s2827 = HState(0) : Shot "DEL_06_2827_after.png"
    Result "DEL-06/REG-04 mixed delivery (2392 ok, 2827 LGORT blank, 1757 dup)", (InStr(s2392, "VBELN_VL=[]") = 0), "per-order success/failure, loop continues", "2392: " & s2392 & " | 2827: " & s2827 & " | log: " & r
    If InStr(s2392, "VBELN_VL=[]") = 0 Then
        Dim vl2392 : vl2392 = Mid(s2392, InStr(s2392, "VBELN_VL=[") + 10, 8)
        Se16n "VBUP", K_DOC, vl2392, "VBELN,POSNR,WBSTA,LFSTA,FKSTA", "2392 VBUP " & vl2392, 12, "DIV_01_2392_vbup.png"
    End If
Else
    log.WriteLine "mixed query failed: screen=" & session.Info.ScreenNumber & " sbar=" & Sbar()
End If

' ---------- ALV-08 : layout popups ----------
If OpenByID(ID_1722) Then
    Dim g : Set g = session.findById(HG)
    On Error Resume Next
    g.pressToolbarContextButton "&MB_VARIANT"
    WScript.Sleep 600
    g.selectContextMenuItem "&LOAD"
    WScript.Sleep 1500
    log.WriteLine "ALV-08 saved layouts popup=[" & WT(1) & "]"
    ShotW 1, "ALV_08_saved_layouts.png"
    CancelAll
    Set g = session.findById(HG)
    g.pressToolbarContextButton "&MB_VARIANT"
    WScript.Sleep 600
    g.selectContextMenuItem "&COL0"
    WScript.Sleep 1500
    log.WriteLine "ALV-08 change layout popup=[" & WT(1) & "]"
    ShotW 1, "ALV_08_change_layout.png"
    CancelAll
    Err.Clear
    On Error GoTo 0
    Note "ALV-08", "layout popups captured; check GI_ICON in displayed/hidden column lists in the shots"
End If
Finish
