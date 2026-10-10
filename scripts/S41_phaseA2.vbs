' S41_phaseA2.vbs : 1757 (now STATU=B) - MAT-01/UNIT-01 (robust entry) -> ATP-06/01 dbl-click -> CREA_ORD -> re-query -> Z33SDT644 -> MAT-06
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i, ci, nPass, nFail
nPass = 0 : nFail = 0
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For ci = 0 To application.Children.Count - 1
    If application.Children.ElementAt(ci).Children.ElementAt(0).Info.SystemName = "KDQ" Then Set connection = application.Children.ElementAt(ci) : Exit For
Next
If connection Is Nothing Then WScript.Echo "NO KDQ" : WScript.Quit 99
Set base = connection.Children.ElementAt(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S41_phaseA2.txt", True, True)
n0 = connection.Children.Count
base.createSession
For i = 1 To 120
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
If connection.Children.Count <= n0 Then WScript.Echo "NO NEW SESSION" : WScript.Quit 98
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
If session.Id = base.Id Then WScript.Echo "REFUSE ses0" : WScript.Quit 98
WScript.Sleep 2500
log.WriteLine "System=" & session.Info.SystemName & " User=" & session.Info.User & " RunAt=" & Now
Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"
Const LOGGRID = "wnd[1]/usr/cntlGRID1/shellcont/shell"
Const SHOTS = "C:\sap-test\results\shots\gi\"
Const ID_1757 = "03f32fd8-6329-11f1-9ba3-000c298aeffe"
Const NEWMAT = "PNGB3500013L008"
Const MATFLD = "wnd[1]/usr/txtGS_SCREEN_0110-MATNR"
Dim K_CHGTITLE, K_YES, K_NO, K_CANCEL
K_CHGTITLE = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HCF54)&ChrW(&HB4DC)
K_YES = ChrW(&HC608) : K_NO = ChrW(&HC544)&ChrW(&HB2C8)&ChrW(&HC624) : K_CANCEL = ChrW(&HCDE8)&ChrW(&HC18C)

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
' close popups top-down: No -> Cancel -> tbar btn[12] -> close  (never Enter)
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
Function YesAndLog(shotName)
    Dim g, s, k, grid, cols, r, cc, ln
    s = ""
    On Error Resume Next
    For g = 1 To 4
        If Not HasW(1) Then Exit For
        If PressBtn("wnd[1]", K_YES) Then WScript.Sleep 2500 : s = s & "[YES] " Else Exit For
    Next
    For k = 1 To 20
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
            If r > 8 Then Exit For
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
Function OpenByID(idv)
    Tcode "/nZSD600"
    Dim w, okSel : okSel = False
    Dim selObj
    For w = 1 To 30
        On Error Resume Next
        Set selObj = session.findById("wnd[0]/usr/radP_RBLIST")
        If Err.Number = 0 Then okSel = True
        Err.Clear
        On Error GoTo 0
        If okSel Then Exit For
        WScript.Sleep 500
    Next
    If Not okSel Then log.WriteLine "  (selection screen not reached: screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " w1=" & WT(1) & " sbar=" & Sbar() & ")" : OpenByID = False : Exit Function
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
Function H(col) : H = session.findById(HG).GetCellValue(0, col) : End Function
Function It(r, col) : It = session.findById(IG).GetCellValue(r, col) : End Function
Sub SelHeader0()
    Dim gh : Set gh = session.findById(HG)
    gh.setCurrentCell 0,"BSTKD" : gh.selectedRows = "0"
End Sub
Function DblHeader0()
    SelHeader0
    On Error Resume Next
    session.findById(HG).doubleClickCurrentCell
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 2000
    DblHeader0 = "sbar[" & SbarT() & ":" & Sbar() & "]"
    If HasW(2) Then DblHeader0 = DblHeader0 & " w2=[" & W2Msg() & "]"
    CancelAll
End Function
Sub SelItem(r)
    session.findById(IG).setCurrentCell r,"MATNR" : session.findById(IG).selectedRows = CStr(r)
End Sub
Function HeaderState()
    HeaderState = "STATU=" & H("STATU") & " VBELN=[" & H("VBELN") & "] VBELN_VL=[" & H("VBELN_VL") & "] GI_ICON=[" & H("GI_ICON") & "]"
End Function
Sub Se16n644(idv, label)
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "Z33SDT644"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    Dim rr, desc
    For rr = 0 To 12
        desc = "" : desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, "Uniq") > 0 Then session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = idv : Exit For
    Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2000
    Dim g, r2
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then log.WriteLine "  [" & label & "] Z33SDT644 -> " & Sbar() : Err.Clear : On Error GoTo 0 : Exit Sub
    log.WriteLine "  [" & label & "] Z33SDT644 rows=" & g.RowCount
    For r2 = 0 To g.RowCount - 1
        log.WriteLine "     SEQNO=" & g.GetCellValue(r2,"SEQNO") & " MATNR_OLD=" & g.GetCellValue(r2,"MATNR_OLD") & " MATNR=" & g.GetCellValue(r2,"MATNR") & " VBELN=" & g.GetCellValue(r2,"VBELN") & " ERNAM=" & g.GetCellValue(r2,"ERNAM") & " ERDAT=" & g.GetCellValue(r2,"ERDAT") & " ERZET=" & g.GetCellValue(r2,"ERZET")
    Next
    Shot label & "_z33sdt644.png"
    On Error GoTo 0
End Sub

Dim r, msg, before, after, vb
If Not OpenByID(ID_1757) Then log.WriteLine "ABORT" : log.Close : WScript.Quit 10
log.WriteLine "1757 start: " & HeaderState()
If H("VBELN") <> "" Or H("STATU") = "C" Then log.WriteLine "ABORT precondition (B, no VBELN)" : log.Close : WScript.Quit 10
msg = DblHeader0()
before = It(0,"MATNR") & "/" & It(0,"MAKTX") & "/" & It(0,"MEINS")
log.WriteLine "  item0 before=" & before & " " & msg

' ---- MAT-01 / UNIT-01 (robust entry) ----
SelItem 0
session.findById(IG).pressToolbarButton "CHANGE_MAT"
WScript.Sleep 1800
If InStr(WT(1), K_CHGTITLE) = 0 Then
    Result "MAT-01", False, "dialog", "w1=" & WT(1) & " sbar=" & Sbar()
    CancelAll
Else
    Dim f, tries, got
    got = ""
    For tries = 1 To 3
        On Error Resume Next
        Set f = session.findById(MATFLD)
        f.SetFocus
        f.Text = NEWMAT
        Err.Clear
        got = session.findById(MATFLD).Text
        On Error GoTo 0
        If got = NEWMAT Then Exit For
        WScript.Sleep 700
    Next
    log.WriteLine "  entry tries=" & tries & " field=[" & got & "]"
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 1500
    Dim dmak : dmak = ""
    On Error Resume Next
    dmak = session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MAKTX").Text
    On Error GoTo 0
    r = "field=[" & got & "] dialogMAKTX=[" & dmak & "] w2=[" & W2Msg() & "]"
    ShotW 1, "MAT_01_dialog.png"
    If HasW(2) Then
        r = r & " (error popup - not saving)"
        CancelAll
    Else
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        WScript.Sleep 1800
        Dim ys : ys = ""
        Do While HasW(1)
            If PressBtn("wnd[1]", K_YES) Then ys = ys & "[YES]" : WScript.Sleep 1400 Else Exit Do
        Loop
        r = r & " " & ys & " sbar[" & SbarT() & ":" & Sbar() & "] w2=[" & W2Msg() & "]"
        CancelAll
    End If
    after = It(0,"MATNR") & "/" & It(0,"MAKTX") & "/" & It(0,"MEINS")
    Shot "MAT_01_after_change.png"
    Result "MAT-01/UNIT-01 material change -> " & NEWMAT & " (in-place ALV)", (It(0,"MATNR") = NEWMAT And It(0,"MEINS") <> ""), "MATNR/MAKTX/MEINS updated", "before=" & before & " after=" & after & " " & r
End If

' ---- ATP-06 / ATP-01 ----
msg = DblHeader0()
Shot "ATP_01_1757_after_dblclick.png"
Result "ATP-06/ATP-01 1757 ATP re-check after change (dbl-click)", (InStr(msg, "ATP") = 0), "no shortage/error message, normal colour (see shot)", msg & " item0=" & It(0,"MATNR") & "/" & It(0,"MEINS")

' ---- CREA_ORD ----
SelHeader0
session.findById(HG).pressToolbarButton "CREA_ORD"
WScript.Sleep 2500
r = "confirm=[" & WT(1) & "] " & YesAndLog("A03_1757_crea_ord_log.png")
Shot "A03_1757_crea_ord_after.png"
vb = ""
If OpenByID(ID_1757) Then vb = H("VBELN") : r = r & " requery: " & HeaderState()
Shot "A04_1757_after_crea_ord.png"
Result "UNIT-01/ATP-01 1757 CREA_ORD -> sales order (no unit/ATP error)", (vb <> ""), "VBELN created", r
msg = DblHeader0()
log.WriteLine "  1757 re-query item0=" & It(0,"MATNR") & "/" & It(0,"MAKTX") & "/" & It(0,"MEINS") & " " & msg
Shot "A05_1757_requery_items.png"
Result "HIST-03(1757) changed material shown after re-query", (It(0,"MATNR") = NEWMAT), "item0 MATNR=" & NEWMAT & " after fresh query (Z33SDT644 overlay)", "item0=" & It(0,"MATNR") & " VBELN=" & vb
Se16n644 ID_1757, "HIST_1757"

' ---- MAT-06 ----
If vb <> "" Then
    If OpenByID(ID_1757) Then
        msg = DblHeader0()
        SelItem 0
        session.findById(IG).pressToolbarButton "CHANGE_MAT"
        WScript.Sleep 1400
        Dim o6 : o6 = (InStr(WT(1), K_CHGTITLE) > 0)
        r = "dialogOpened=" & o6 & " sbar[" & SbarT() & ":" & Sbar() & "]"
        Shot "MAT_06_after_so_blocked.png"
        CancelAll
        Result "MAT-06 change after sales order (VBELN " & vb & ")", (Not o6), "blocked with sales-order message", r
    End If
End If

Tcode "/n"
log.WriteLine "SUMMARY: PASS=" & nPass & " FAIL=" & nFail
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "S41 done PASS=" & nPass & " FAIL=" & nFail
