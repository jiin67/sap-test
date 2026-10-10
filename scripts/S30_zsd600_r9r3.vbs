' S30_zsd600_r9r3.vbs : ZSD600 modified-program test - R9 material-change guards + R3/DEL delivery guards
'   + R9-01 designated positive (change ML00007621->PNGB3500035L001 on BSTKD 2190, then REVERT).
'   Safety: confirm dialogs answered No/Cancel only; never Yes/Save except the R9-01 designated change,
'           which is reverted at the end. Source is pure ASCII; Korean matched via ChrW.
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Dim nPass, nFail
nPass = 0 : nFail = 0
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S30_zsd600_r9r3_" & DateStamp() & ".txt", True, True)
If connection.Children.Count > 1 Then
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
Else
    n0 = connection.Children.Count
    base.createSession
    For i = 1 To 40
        WScript.Sleep 500
        If connection.Children.Count > n0 Then Exit For
    Next
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    WScript.Sleep 1500
End If
log.WriteLine "System=" & session.Info.SystemName & " Client=" & session.Info.Client & " User=" & session.Info.User & " RunAt=" & Now

Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"

' Korean match strings via ChrW (space-safe, single contiguous words)
Dim K_ONEITEM, K_INPUTMAT, K_NOTINMASTER, K_NOTINPLANT, K_SAME, K_DELIV, K_SELECT, K_CHGMATTITLE, K_NO, K_CANCEL, K_CLOSED, K_HASSO, K_SAVEQ
K_ONEITEM     = ChrW(&HAC74) & ChrW(&HB9CC)                                   ' gun-man
K_INPUTMAT    = ChrW(&HC790) & ChrW(&HC7AC) & ChrW(&HBC88) & ChrW(&HD638) & ChrW(&HB97C)   ' jaejae-beonho-reul
K_NOTINMASTER = ChrW(&HC790) & ChrW(&HC7AC) & ChrW(&HB9C8) & ChrW(&HC2A4) & ChrW(&HD130)   ' jaejae-master
K_NOTINPLANT  = ChrW(&HD50C) & ChrW(&HB79C) & ChrW(&HD2B8)                    ' plant
K_SAME        = ChrW(&HAC19) & ChrW(&HC2B5) & ChrW(&HB2C8) & ChrW(&HB2E4)     ' gat-seumnida
K_DELIV       = ChrW(&HB0A9) & ChrW(&HD488)                                    ' nappum
K_SELECT      = ChrW(&HC120) & ChrW(&HD0DD)                                    ' seontaek
K_CHGMATTITLE = ChrW(&HC790) & ChrW(&HC7AC) & ChrW(&HCF54) & ChrW(&HB4DC)     ' jaejae-code
K_NO          = ChrW(&HC544) & ChrW(&HB2C8) & ChrW(&HC624)                    ' aniyo
K_CANCEL      = ChrW(&HCDE8) & ChrW(&HC18C)                                    ' chwiso
K_CLOSED      = ChrW(&HB9C8) & ChrW(&HAC10)                                    ' magam (closed)
K_HASSO       = ChrW(&HD310) & ChrW(&HB9E4)                                    ' panmae (sales)
K_SAVEQ       = ChrW(&HC800) & ChrW(&HC7A5)                                    ' jeojang (save)

Function DateStamp()
    Dim d : d = Now
    DateStamp = Year(d) & Right("0" & Month(d),2) & Right("0" & Day(d),2) & "_" & Right("0" & Hour(d),2) & Right("0" & Minute(d),2)
End Function
Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Function Sbar()
    Sbar = session.findById("wnd[0]/sbar").Text
End Function
Function W1Text(s_)
    On Error Resume Next
    W1Text = ""
    W1Text = session.findById("wnd[1]" & s_).Text
    Err.Clear
    On Error GoTo 0
End Function
Function HasPopup()
    On Error Resume Next
    HasPopup = False
    Dim w : Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then HasPopup = True
    Err.Clear
    On Error GoTo 0
End Function
Function PopupText()
    Dim usr, c, comp, s
    s = ""
    On Error Resume Next
    Set usr = session.findById("wnd[1]/usr")
    If Err.Number <> 0 Then Err.Clear : PopupText = "" : Exit Function
    On Error GoTo 0
    For c = 0 To usr.Children.Count - 1
        Set comp = usr.Children.ElementAt(c)
        On Error Resume Next
        If comp.Type = "GuiLabel" Or comp.Type = "GuiTextField" Then s = s & "|" & comp.Text
        On Error GoTo 0
    Next
    PopupText = s
End Function
' Safe-close: No -> Cancel -> btn[12] -> close. NEVER Yes/Save.
Sub CancelPopups()
    Dim g, wname, usr, c, comp, pressed, w2
    On Error Resume Next
    For g = 1 To 8
        If Not HasPopup() Then Exit Sub
        wname = "wnd[1]"
        Set w2 = session.findById("wnd[2]")
        If Err.Number = 0 Then wname = "wnd[2]"
        Err.Clear
        pressed = False
        Set usr = session.findById(wname & "/usr")
        If Err.Number = 0 Then
            For c = 0 To usr.Children.Count - 1
                Set comp = usr.Children.ElementAt(c)
                If comp.Type = "GuiButton" Then
                    If InStr(comp.Text & comp.Tooltip, K_NO) > 0 Or InStr(comp.Text & comp.Tooltip, K_CANCEL) > 0 _
                       Or InStr(LCase(comp.Text & comp.Tooltip), "cancel") > 0 Or InStr(LCase(comp.Text & comp.Tooltip), "no") > 0 Then
                        comp.press : pressed = True : Exit For
                    End If
                End If
            Next
        End If
        Err.Clear
        If Not pressed Then session.findById(wname & "/tbar[0]/btn[12]").press : Err.Clear
        If HasPopup() Then session.findById(wname).close : Err.Clear
        WScript.Sleep 700
    Next
    On Error GoTo 0
End Sub
Sub Result(tc, ok, expected, actual)
    If ok Then nPass = nPass+1 : log.WriteLine "[PASS] " & tc Else nFail = nFail+1 : log.WriteLine "[FAIL] " & tc
    log.WriteLine "       expect: " & expected
    log.WriteLine "       actual: " & actual
End Sub
' query one BSTKD, all status, wide date -> screen 100?
Function OpenBSTKD(bstkd)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = bstkd
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 3500
    OpenBSTKD = (session.Info.ScreenNumber = 100)
End Function
' query by status only (find candidate) -> returns first BSTKD or ""
Function FirstBstkdByStatus(st)
    FirstBstkdByStatus = ""
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = st
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = st
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 4000
    If session.Info.ScreenNumber = 100 Then
        On Error Resume Next
        FirstBstkdByStatus = session.findById(HG).GetCellValue(0, "BSTKD")
        Err.Clear
        On Error GoTo 0
    End If
End Function
Sub SelHeader(r)
    Dim gh : Set gh = session.findById(HG)
    gh.setCurrentCell r, "BSTKD"
    gh.selectedRows = CStr(r)
    On Error Resume Next
    gh.doubleClickCurrentCell
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 1500
    CancelPopups
End Sub
Function ItemRows()
    On Error Resume Next
    ItemRows = 0
    ItemRows = session.findById(IG).RowCount
    Err.Clear
    On Error GoTo 0
End Function
Function HeaderRows()
    On Error Resume Next
    HeaderRows = 0
    HeaderRows = session.findById(HG).RowCount
    Err.Clear
    On Error GoTo 0
End Function
Function OpenChangeMat()
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1500
    OpenChangeMat = (InStr(W1Text(""), K_CHGMATTITLE) > 0)
End Function
' set material and Save; returns state string; caller cancels after. saveConfirmYes=True only for designated R9-01.
Function ChangeMatSave(matnr, saveConfirmYes)
    Dim r, e1, e2
    On Error Resume Next
    session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = matnr
    e1 = Err.Number : Err.Clear
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 800
    session.findById("wnd[1]/tbar[0]/btn[8]").press   ' Save
    e2 = Err.Number : Err.Clear
    On Error GoTo 0
    WScript.Sleep 1600
    ' a save-confirm (Yes/No) may appear
    If saveConfirmYes And HasPopup() Then
        Dim usr, c, comp
        On Error Resume Next
        Set usr = session.findById("wnd[1]/usr")
        For c = 0 To usr.Children.Count - 1
            Set comp = usr.Children.ElementAt(c)
            If comp.Type = "GuiButton" Then
                If InStr(comp.Text & comp.Tooltip, ChrW(&HC608)) > 0 Then comp.press : Exit For   ' 예 (Yes) - ONLY for designated case
            End If
        Next
        Err.Clear
        On Error GoTo 0
        WScript.Sleep 1600
    End If
    r = "setErr=" & e1 & " saveErr=" & e2 & " sbar[" & session.findById("wnd[0]/sbar").MessageType & ":" & Sbar() & "]"
    r = r & " w1=[" & W1Text("") & "]"
    On Error Resume Next
    Dim w2t : w2t = "" : w2t = session.findById("wnd[2]").Text
    If Err.Number = 0 Then r = r & " wnd2=[" & w2t & "]"
    Err.Clear
    On Error GoTo 0
    If Not HasPopup() Then r = r & " (dialog closed screen=" & session.Info.ScreenNumber & ")"
    ChangeMatSave = r
End Function

' ======================= R9 guards =======================
Sub R9_02()
    Dim tc : tc = "R9-02 no item selected"
    If Not OpenBSTKD("2190") Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    SelHeader 0
    If ItemRows() = 0 Then Result tc, False, "items>0", "0 items" : Exit Sub
    session.findById(IG).selectedRows = ""
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1200
    Dim pt : pt = W1Text("") : Dim sb : sb = Sbar()
    CancelPopups
    Result tc, (pt = "" And InStr(sb, K_ONEITEM) > 0), "M20 one-item guard, no dialog", "sbar=[" & sb & "] popup=[" & pt & "]"
End Sub
Sub R9_03(multiBstkd)
    Dim tc : tc = "R9-03 multiple items selected"
    If multiBstkd = "" Then Result tc, False, "an order with >=2 items", "no multi-item order found (needs test data)" : Exit Sub
    If Not OpenBSTKD(multiBstkd) Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    SelHeader 0
    If ItemRows() < 2 Then Result tc, False, "items>=2", "items=" & ItemRows() & " (bstkd " & multiBstkd & ")" : Exit Sub
    session.findById(IG).selectedRows = "0,1"
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1200
    Dim pt : pt = W1Text("") : Dim sb : sb = Sbar()
    CancelPopups
    Result tc, (pt = "" And InStr(sb, K_ONEITEM) > 0), "one-item-only message, no dialog", "sbar=[" & sb & "] popup=[" & pt & "]"
End Sub
Sub R9_04()
    Dim tc : tc = "R9-04 same material entered"
    If Not OpenBSTKD("2190") Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    SelHeader 0
    If ItemRows() = 0 Then Result tc, False, "items>0", "0 items" : Exit Sub
    Dim old : old = session.findById(IG).GetCellValue(0, "MATNR")
    session.findById(IG).setCurrentCell 0, "MATNR" : session.findById(IG).selectedRows = "0"
    If Not OpenChangeMat() Then Result tc, False, "dialog", "w1=" & W1Text("") : CancelPopups : Exit Sub
    Dim r : r = ChangeMatSave(old, False)
    CancelPopups
    Result tc, (InStr(r,K_SAME)>0 Or InStr(r,"dialog closed")>0), "M25 same-material blocked or no-op (old=" & old & ")", r
End Sub
Sub R9_05()
    Dim tc : tc = "R9-05 unregistered material (MARA missing)"
    If Not OpenBSTKD("2190") Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    SelHeader 0
    If ItemRows() = 0 Then Result tc, False, "items>0", "0 items" : Exit Sub
    session.findById(IG).setCurrentCell 0, "MATNR" : session.findById(IG).selectedRows = "0"
    If Not OpenChangeMat() Then Result tc, False, "dialog", "w1=" & W1Text("") : CancelPopups : Exit Sub
    Dim r : r = ChangeMatSave("ZZNOEXIST9", False)
    Dim rejected : rejected = (InStr(r,"wnd2=[")>0 Or InStr(r,K_NOTINMASTER)>0 Or InStr(r,"dialog closed")=0)
    CancelPopups
    Result tc, rejected, "M23 not-in-master: error popup or dialog stays (no save)", r
End Sub
Sub R9_07(vbelnBstkd)
    Dim tc : tc = "R9-07 change on order with sales doc (VBELN)"
    If vbelnBstkd = "" Then Result tc, False, "order with VBELN", "no VBELN-bearing order found (needs test data)" : Exit Sub
    If Not OpenBSTKD(vbelnBstkd) Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    Dim vb : vb = session.findById(HG).GetCellValue(0, "VBELN")
    If vb = "" Then Result tc, False, "VBELN present", "bstkd " & vbelnBstkd & " VBELN empty" : Exit Sub
    SelHeader 0
    If ItemRows() = 0 Then Result tc, False, "items>0", "0 items" : Exit Sub
    session.findById(IG).setCurrentCell 0, "MATNR" : session.findById(IG).selectedRows = "0"
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1300
    Dim opened : opened = (InStr(W1Text(""), K_CHGMATTITLE) > 0)
    Dim sb : sb = Sbar()
    CancelPopups
    Result tc, (Not opened), "change blocked (VBELN=" & vb & "): dialog does NOT open", "dialogOpened=" & opened & " sbar=[" & sb & "]"
End Sub
Sub R9_08(closedBstkd)
    Dim tc : tc = "R9-08 change on closed order (STATU=C)"
    If closedBstkd = "" Then Result tc, False, "closed order STATU=C", "no closed order found (needs test data)" : Exit Sub
    If Not OpenBSTKD(closedBstkd) Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    SelHeader 0
    If ItemRows() = 0 Then Result tc, False, "items>0", "0 items" : Exit Sub
    session.findById(IG).setCurrentCell 0, "MATNR" : session.findById(IG).selectedRows = "0"
    session.findById(IG).pressToolbarButton "CHANGE_MAT"
    WScript.Sleep 1300
    Dim opened : opened = (InStr(W1Text(""), K_CHGMATTITLE) > 0)
    Dim sb : sb = Sbar()
    CancelPopups
    Result tc, (Not opened), "change blocked on closed order: dialog does NOT open", "dialogOpened=" & opened & " sbar=[" & sb & "]"
End Sub

' ======================= R3 / DEL guards =======================
Sub R3_02()
    Dim tc : tc = "R3-02 create delivery on order without sales doc"
    If Not OpenBSTKD("2190") Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    Dim gh : Set gh = session.findById(HG)
    If gh.GetCellValue(0,"VBELN") <> "" Then Result tc, False, "VBELN empty", "VBELN=" & gh.GetCellValue(0,"VBELN") : Exit Sub
    gh.setCurrentCell 0,"BSTKD" : gh.selectedRows = "0"
    gh.pressToolbarButton "CRE_DELIVERY"
    WScript.Sleep 2500
    Dim pt : pt = W1Text("") : Dim ptx : ptx = PopupText() : Dim sb : sb = Sbar()
    CancelPopups
    WScript.Sleep 600
    Dim vl : vl = "?"
    If OpenBSTKD("2190") Then vl = session.findById(HG).GetCellValue(0,"VBELN_VL")
    Result tc, (vl = ""), "blocked / No-cancel; VBELN_VL stays empty", "confirm=[" & pt & " :: " & ptx & "] sbar=[" & sb & "] after VBELN_VL=[" & vl & "]"
End Sub
Sub DEL_04()
    Dim tc : tc = "DEL-04 delete delivery when none exists"
    If Not OpenBSTKD("2190") Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    Dim gh : Set gh = session.findById(HG)
    If gh.GetCellValue(0,"VBELN_VL") <> "" Then Result tc, False, "no delivery precond", "VBELN_VL=" & gh.GetCellValue(0,"VBELN_VL") : Exit Sub
    gh.setCurrentCell 0,"BSTKD" : gh.selectedRows = "0"
    gh.pressToolbarButton "DEL_DELIVERY"
    WScript.Sleep 2200
    Dim pt : pt = W1Text("") : Dim ptx : ptx = PopupText() : Dim sb : sb = Sbar()
    CancelPopups
    Result tc, (InStr(sb & ptx, K_DELIV) > 0), "no-delivery-to-delete message", "sbar=[" & sb & "] popup=[" & pt & " :: " & ptx & "]"
End Sub

' ======================= R9-01 designated positive (change + REVERT) =======================
Sub R9_01()
    Dim tc : tc = "R9-01 material change save (ML00007621 -> PNGB3500035L001) then revert"
    If Not OpenBSTKD("2190") Then Result tc, False, "screen100", "screen=" & session.Info.ScreenNumber : Exit Sub
    ' pick the header whose item is ML00007621 and VBELN empty and STATU<>C
    Dim gh : Set gh = session.findById(HG)
    Dim hr, useR, found
    found = False
    For hr = 0 To gh.RowCount - 1
        If gh.GetCellValue(hr,"VBELN") = "" And gh.GetCellValue(hr,"STATU") <> "C" Then
            SelHeader hr
            If ItemRows() >= 1 Then
                If session.findById(IG).GetCellValue(0,"MATNR") = "ML00007621" Then useR = hr : found = True : Exit For
            End If
        End If
    Next
    If Not found Then Result tc, False, "header with ML00007621, VBELN empty, not closed", "not found among 2190 headers" : Exit Sub
    Dim before : before = session.findById(IG).GetCellValue(0,"MATNR")
    log.WriteLine "  R9-01 target header row=" & useR & " item MATNR(before)=" & before
    ' change to new material and save (designated) -> answer Yes on save-confirm if any
    session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
    If Not OpenChangeMat() Then Result tc, False, "dialog", "w1=" & W1Text("") : CancelPopups : Exit Sub
    Dim r : r = ChangeMatSave("PNGB3500035L001", True)
    log.WriteLine "  R9-01 save result: " & r
    CancelPopups
    ' re-read to see if applied
    Dim after : after = "?"
    If OpenBSTKD("2190") Then
        SelHeader useR
        If ItemRows() >= 1 Then after = session.findById(IG).GetCellValue(0,"MATNR")
    End If
    Dim changed : changed = (after = "PNGB3500035L001")
    log.WriteLine "  R9-01 item MATNR(after)=" & after & " changed=" & changed
    ' ---- REVERT back to ML00007621 to clean up ----
    Dim reverted : reverted = False
    If changed Then
        SelHeader useR
        session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
        If OpenChangeMat() Then
            Dim r2 : r2 = ChangeMatSave("ML00007621", True)
            log.WriteLine "  R9-01 revert result: " & r2
            CancelPopups
            If OpenBSTKD("2190") Then
                SelHeader useR
                If ItemRows() >= 1 Then reverted = (session.findById(IG).GetCellValue(0,"MATNR") = "ML00007621")
            End If
        End If
    End If
    log.WriteLine "  R9-01 reverted=" & reverted
    Result tc, changed, "material change applies (MATNR/MAKTX/MEINS updated); reverted afterwards=" & reverted, _
        "before=" & before & " after=" & after & " reverted=" & reverted
End Sub

' ---- find candidates ----
Dim closedBstkd, vbelnBstkd, multiBstkd
closedBstkd = FirstBstkdByStatus("C")
log.WriteLine "candidate closed(STATU=C) BSTKD=" & closedBstkd
' VBELN-bearing candidate: try known order from earlier session
vbelnBstkd = ""
If OpenBSTKD("1422 " & ChrW(&H43E) & ChrW(&H442) & " 12.05.2026") Then
    If session.findById(HG).GetCellValue(0,"VBELN") <> "" Then vbelnBstkd = "1422 " & ChrW(&H43E) & ChrW(&H442) & " 12.05.2026"
End If
log.WriteLine "candidate VBELN BSTKD=[" & vbelnBstkd & "]"
' multi-item candidate: scan a busy month
multiBstkd = ""
Tcode "/nZSD600"
With session
    .findById("wnd[0]/usr/radP_RBLIST").Select
    .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.08.01"
    .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.08.31"
    .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = "B"
    .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = "B"
    .findById("wnd[0]/tbar[1]/btn[8]").press
End With
WScript.Sleep 5000
If session.Info.ScreenNumber = 100 Then
    Dim hh, hn
    hn = HeaderRows()
    If hn > 12 Then hn = 12
    For hh = 0 To hn - 1
        SelHeader hh
        If ItemRows() >= 2 Then multiBstkd = session.findById(HG).GetCellValue(hh,"BSTKD") : Exit For
    Next
End If
log.WriteLine "candidate multi-item BSTKD=[" & multiBstkd & "]"

' ---- run ----
On Error Resume Next
R9_02 : If Err.Number<>0 Then Result "R9-02",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
R9_03 multiBstkd : If Err.Number<>0 Then Result "R9-03",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
R9_04 : If Err.Number<>0 Then Result "R9-04",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
R9_05 : If Err.Number<>0 Then Result "R9-05",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
R9_07 vbelnBstkd : If Err.Number<>0 Then Result "R9-07",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
R9_08 closedBstkd : If Err.Number<>0 Then Result "R9-08",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
R3_02 : If Err.Number<>0 Then Result "R3-02",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
DEL_04 : If Err.Number<>0 Then Result "DEL-04",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
R9_01 : If Err.Number<>0 Then Result "R9-01",False,"-","runtime " & Err.Description : Err.Clear
CancelPopups
On Error GoTo 0

Tcode "/n"
log.WriteLine "SUMMARY: PASS=" & nPass & " FAIL=" & nFail
log.Close
If Not HasPopup() Then session.findById("wnd[0]").close
WScript.Echo "S30 done PASS=" & nPass & " FAIL=" & nFail
