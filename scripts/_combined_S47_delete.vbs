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

' S47 body : delivery deletion + GI cancel lifecycle, history verify, header de-dup
'   DELETE-01/ALV-07 : 1757 delivery 85857217 (WBSTA A) -> DEL_DELIVERY -> deleted, VBELN_VL & GI blank
'   DELETE-02       : 1722 delivery 85857218 (WBSTA C, GI'd) -> DEL_DELIVERY -> GI cancel + delete, VBUP back to A
'   DELETE-05       : 1757 (now no delivery) -> DEL_DELIVERY -> 'no delivery' message
'   2827 delivery 85857219 -> delete (cleanup)
'   HIST-02/03 2318 : SO 4293339 exists (item11 changed) -> Z33SDT644 + re-query
'   REG-01          : header de-dup by material query
Const ID_1757 = "03f32fd8-6329-11f1-9ba3-000c298aeffe"
Const ID_1722 = "04a27fe5-6007-11f1-9ba0-000c298aeffe"
Const ID_2827 = "04bf872f-a20b-11f1-9bce-000c298aeffe"
Const ID_2318 = "0407e0e8-8434-11f1-9bbb-000c298aeffe"
Dim r, vl, gi

' DELETE-01 : 1757 WBSTA A -> delete
If OpenByID(ID_1757) Then
    vl = Hr(0,"VBELN_VL")
    log.WriteLine "1757 before delete: " & HState(0)
    If vl <> "" Then
        r = HeaderAction(0, "DEL_DELIVERY", "DELETE_01_1757_log.png")
        Shot "DELETE_01_1757_after_inplace.png"
        Dim vlIn : vlIn = Hr(0,"VBELN_VL")
        If OpenByID(ID_1757) Then r = r & " requery: " & HState(0)
        Shot "DELETE_01_1757_requery.png"
        Result "DELETE-01/ALV-07 1757 unshipped delivery (" & vl & ") deleted, VBELN_VL & GI cleared", (Hr(0,"VBELN_VL") = "" And Hr(0,"GI_ICON") = ""), "VBELN_VL blank, GI blank", "inplaceVL=[" & vlIn & "] " & r
        ' DELETE-05 : now no delivery -> delete again
        r = HeaderAction(0, "DEL_DELIVERY", "DELETE_05_1757_none_log.png")
        Result "DELETE-05 1757 delete with no delivery -> 'none' message", (InStr(r, K_DOC) > 0 Or InStr(r, ChrW(&HC5C6)&ChrW(&HC2B5)) > 0 Or InStr(r, "RESULT=[E]") > 0), "no-delivery message", r
    Else
        Note "DELETE-01 1757", "no delivery present: " & HState(0)
    End If
End If

' DELETE-02 : 1722 WBSTA C (GI'd) -> GI cancel + delete
If OpenByID(ID_1722) Then
    vl = Hr(0,"VBELN_VL") : gi = Hr(0,"GI_ICON")
    log.WriteLine "1722 before delete: " & HState(0)
    If vl <> "" Then
        r = HeaderAction(0, "DEL_DELIVERY", "DELETE_02_1722_log.png")
        Shot "DELETE_02_1722_after_inplace.png"
        Dim vlIn2, giIn2 : vlIn2 = Hr(0,"VBELN_VL") : giIn2 = Hr(0,"GI_ICON")
        If OpenByID(ID_1722) Then r = r & " requery: " & HState(0)
        Shot "DELETE_02_1722_requery.png"
        Result "DELETE-02 1722 GI-completed delivery (" & vl & ", GI " & gi & ") -> GI cancel + delete", (Hr(0,"VBELN_VL") = "" And Hr(0,"GI_ICON") = ""), "GI cancelled, delivery deleted, VBELN_VL & GI blank", "beforeGI=[" & gi & "] inplaceVL=[" & vlIn2 & "] inplaceGI=[" & giIn2 & "] " & r
    End If
End If

' 2827 delivery cleanup
If OpenByID(ID_2827) Then
    vl = Hr(0,"VBELN_VL")
    If vl <> "" Then
        r = HeaderAction(0, "DEL_DELIVERY", "S47_2827_del_log.png")
        If OpenByID(ID_2827) Then r = r & " after: " & HState(0)
        Note "cleanup 2827 delivery " & vl & " deleted", r
    End If
End If

' HIST-02/03 : 2318 (SO 4293339, item11 changed earlier)
If OpenByID(ID_2318) Then
    log.WriteLine "2318: " & HState(0)
    Se16n "Z33SDT644", "Uniq", ID_2318, "SEQNO,MATNR_OLD,MATNR,VBELN,ERNAM,ERDAT,ERZET", "HIST-02 Z33SDT644 2318", 12, "HIST_02_2318_z33sdt644.png"
    Dim msg : msg = DblHeader(0)
    log.WriteLine "  2318 items: item11(SEQNO11)=" & It(10,"MATNR")
    Shot "HIST_03_2318_requery_items.png"
    Result "HIST-02/03 2318 changed material persisted + shown after re-query", (Hr(0,"VBELN") <> "" And It(10,"MATNR") = "PNGB3500013L008"), "Z33SDT644 has row, item11 shows changed material", "VBELN=" & Hr(0,"VBELN") & " item11=" & It(10,"MATNR")
End If

' REG-01 : header de-dup by material
Tcode "/nZSD600"
If WaitSel() Then
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = "PNGB3500013L008"
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 4000
    If session.Info.ScreenNumber = 100 Then
        Dim n, d, dup, rr, key
        n = HeaderRows()
        Set d = CreateObject("Scripting.Dictionary")
        dup = 0
        Dim vis : vis = session.findById(HG).VisibleRowCount
        Dim pg
        For pg = 0 To n - 1 Step vis
            session.findById(HG).FirstVisibleRow = pg
            WScript.Sleep 200
            For rr = pg To pg + vis - 1
                If rr >= n Then Exit For
                key = Hr(rr,"ID") & "|" & Hr(rr,"BSTKD") & "|" & Hr(rr,"VERSION")
                If Hr(rr,"ID") <> "" Then
                    If d.Exists(key) Then dup = dup + 1 Else d.Add key, 1
                End If
            Next
        Next
        Result "REG-01 header de-dup (material PNGB3500013L008 query)", (dup = 0 And n > 0), "each ID|BSTKD|VERSION once", "rows=" & n & " uniqueKeys=" & d.Count & " duplicates=" & dup
    End If
End If
Finish
