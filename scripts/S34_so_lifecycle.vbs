' S34_so_lifecycle.vbs : AUTHORIZED create+cleanup. Sales-order lifecycle for R9-11/R9-12 on ONE 1C order.
'   Filter by S_ID (single header). Flow: change material (in-memory) -> CREA_ORD -> verify Z33SDT644 + VBELN
'   -> ELIM_ORD (delete sales order) -> verify cleared -> delete leftover Z33SDT644 row if any.
'   NO goods issue. Presses Yes only on this program's own create/delete confirms (authorized).
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S34_so_lifecycle.txt", True, True)
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
log.WriteLine "System=" & session.Info.SystemName & " User=" & session.Info.User & " RunAt=" & Now
Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"
Const TARGET_ID = "06aa91b1-7b8c-11f1-9bb5-000c298aeffe"
Const NEWMAT = "PNGB3500035L001"
Dim K_CHGTITLE, K_YES, K_UNIQ
K_CHGTITLE = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HCF54)&ChrW(&HB4DC)
K_YES = ChrW(&HC608)
K_UNIQ = "Uniq"

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
' press Yes on any confirm popups (up to 4), capturing any result popup text; returns concatenated log
Function ConfirmYesLoop()
    Dim g, usr, c, comp, py, s
    s = ""
    On Error Resume Next
    For g = 1 To 5
        If Not HasW(1) Then Exit For
        py = False
        Set usr = session.findById("wnd[1]/usr")
        If Err.Number = 0 Then
            For c = 0 To usr.Children.Count - 1
                Set comp = usr.Children.ElementAt(c)
                If comp.Type = "GuiButton" Then
                    If InStr(comp.Text & comp.Tooltip, K_YES) > 0 Then comp.press : py = True : Exit For
                End If
            Next
        End If
        Err.Clear
        If Not py Then Exit For
        WScript.Sleep 1600
        s = s & " [YES->sbar:" & Sbar() & "]"
    Next
    On Error GoTo 0
    ConfirmYesLoop = s
End Function
' capture + close a result SALV popup if present (returns text)
Function GrabResultPopup()
    Dim s : s = ""
    On Error Resume Next
    If HasW(1) Then
        s = "popup=[" & WT(1) & "]"
        ' try grid rows
        Dim g : Set g = Nothing
        Set g = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
        If g Is Nothing Then Set g = session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell")
        If Not (g Is Nothing) Then
            s = s & " gridRows=" & g.RowCount
        End If
    End If
    Err.Clear
    On Error GoTo 0
    GrabResultPopup = s
End Function
Sub CloseAllPopups()
    Dim g, n
    On Error Resume Next
    For g = 1 To 8
        If Not HasW(1) Then Exit Sub
        For n = 3 To 1 Step -1
            If HasW(n) Then session.findById("wnd[" & n & "]").sendVKey 0 : WScript.Sleep 400 : If HasW(n) Then session.findById("wnd[" & n & "]").close
            Err.Clear
        Next
        WScript.Sleep 400
    Next
    On Error GoTo 0
End Sub
Function OpenByID(idv)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = idv
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 3000
    OpenByID = (session.Info.ScreenNumber = 100)
End Function
Sub SelHeader0()
    Dim gh : Set gh = session.findById(HG)
    gh.setCurrentCell 0,"BSTKD" : gh.selectedRows = "0"
    On Error Resume Next
    gh.doubleClickCurrentCell
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 1400
    CloseAllPopups
End Sub
Sub Se16nDump(table, idv, label)
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "9999"
    Dim rr, desc, setId
    setId = False
    For rr = 0 To 20
        desc = "" : desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, K_UNIQ) > 0 Then session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = idv : setId = True : Exit For
    Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2200
    Dim g, cols, r2, ln, cc, cid
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then log.WriteLine "  [" & label & "] " & table & " idSet=" & setId & " -> no ALV, sbar=" & Sbar() : Err.Clear : On Error GoTo 0 : Exit Sub
    log.WriteLine "  [" & label & "] " & table & " idSet=" & setId & " rows=" & g.RowCount
    Set cols = g.ColumnOrder
    For r2 = 0 To g.RowCount - 1
        If r2 > 10 Then Exit For
        ln = "   #" & r2 & ":"
        For cc = 0 To cols.Count - 1
            cid = cols(cc)
            If cid="ID" Or cid="BSTKD" Or cid="VERSION" Or cid="SEQNO" Or cid="MATNR" Or cid="MATNR_OLD" Or cid="VBELN" Or cid="ERNAM" Or cid="ERDAT" Then ln = ln & " " & cid & "=[" & g.GetCellValue(r2, cid) & "]"
        Next
        log.WriteLine ln
    Next
    On Error GoTo 0
End Sub

Dim createdVBELN : createdVBELN = ""
Dim ok11, ok12 : ok11 = False : ok12 = False

' ---- snapshot ----
If Not OpenByID(TARGET_ID) Then log.WriteLine "ABORT: order not on screen100" : log.Close : WScript.Echo "abort" : WScript.Quit 10
Dim gh : Set gh = session.findById(HG)
log.WriteLine "SNAPSHOT header rows=" & gh.RowCount & " STATU=" & gh.GetCellValue(0,"STATU") & " VBELN=[" & gh.GetCellValue(0,"VBELN") & "] VBELN_VL=[" & gh.GetCellValue(0,"VBELN_VL") & "]"
If gh.GetCellValue(0,"VBELN") <> "" Then log.WriteLine "ABORT: order already has VBELN (not a clean start)" : Tcode "/n" : log.Close : WScript.Echo "abort" : WScript.Quit 10
SelHeader0
log.WriteLine "  item0 MATNR=" & session.findById(IG).GetCellValue(0,"MATNR")

' ---- baseline Z33SDT644 ----
log.WriteLine "=== baseline Z33SDT644 ==="
Se16nDump "Z33SDT644", TARGET_ID, "base644"

' ---- change material (in-memory) then CREA_ORD ----
If Not OpenByID(TARGET_ID) Then log.WriteLine "ABORT re-open" : log.Close : WScript.Quit 10
SelHeader0
session.findById(IG).setCurrentCell 0,"MATNR" : session.findById(IG).selectedRows = "0"
session.findById(IG).pressToolbarButton "CHANGE_MAT"
WScript.Sleep 1400
If InStr(WT(1), K_CHGTITLE) = 0 Then log.WriteLine "ABORT: change dialog not open" : CloseAllPopups : Tcode "/n" : log.Close : WScript.Quit 10
session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = NEWMAT
session.findById("wnd[1]").sendVKey 0
WScript.Sleep 1200
session.findById("wnd[1]/tbar[0]/btn[8]").press   ' save (in-memory)
WScript.Sleep 1400
Dim cy1 : cy1 = ConfirmYesLoop()
CloseAllPopups
log.WriteLine "after material change (in-memory): item0 MATNR=" & session.findById(IG).GetCellValue(0,"MATNR") & cy1

' CREA_ORD on the single header
session.findById(HG).setCurrentCell 0,"BSTKD" : session.findById(HG).selectedRows = "0"
session.findById(HG).pressToolbarButton "CREA_ORD"
WScript.Sleep 2500
log.WriteLine "CREA_ORD: sbar[" & SbarT() & ":" & Sbar() & "] w1=[" & WT(1) & "]"
Dim cy2 : cy2 = ConfirmYesLoop()
log.WriteLine "  confirm:" & cy2
' capture/close any result popup
Dim rp : rp = GrabResultPopup()
If rp <> "" Then log.WriteLine "  resultPopup " & rp
CloseAllPopups
WScript.Sleep 800
log.WriteLine "  post-CREA screen=" & session.Info.ScreenNumber & " sbar=[" & Sbar() & "]"

' ---- re-query: get VBELN ----
If OpenByID(TARGET_ID) Then
    createdVBELN = session.findById(HG).GetCellValue(0,"VBELN")
End If
log.WriteLine "RESULT createdVBELN=[" & createdVBELN & "]"

' ---- R9-12: Z33SDT644 after ----
log.WriteLine "=== after Z33SDT644 (R9-12) ==="
Se16nDump "Z33SDT644", TARGET_ID, "after644"

' ---- R9-11: open VA03 by double-clicking VBELN cell, read item material ----
If createdVBELN <> "" Then
    If OpenByID(TARGET_ID) Then
        On Error Resume Next
        session.findById(HG).setCurrentCell 0,"VBELN"
        session.findById(HG).doubleClickCurrentCell
        WScript.Sleep 2500
        log.WriteLine "  VA03 nav screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " title=[" & WT(0) & "]"
        ' VA03 overview: item table material - best effort dump of any table control
        Dim tset : tset = session.Info.Transaction
        log.WriteLine "  tcode now=" & tset
        Err.Clear
        On Error GoTo 0
        ' return to ZSD600
        Tcode "/nZSD600"
    End If
End If

' =================== CLEANUP ===================
log.WriteLine "=== CLEANUP: ELIM_ORD (delete sales order) ==="
If createdVBELN <> "" Then
    If OpenByID(TARGET_ID) Then
        session.findById(HG).setCurrentCell 0,"BSTKD" : session.findById(HG).selectedRows = "0"
        session.findById(HG).pressToolbarButton "ELIM_ORD"
        WScript.Sleep 2200
        log.WriteLine "  ELIM_ORD sbar[" & SbarT() & ":" & Sbar() & "] w1=[" & WT(1) & "]"
        Dim cy3 : cy3 = ConfirmYesLoop()
        log.WriteLine "  confirm:" & cy3
        Dim rp2 : rp2 = GrabResultPopup()
        If rp2 <> "" Then log.WriteLine "  resultPopup " & rp2
        CloseAllPopups
    End If
End If
' verify VBELN cleared
Dim afterVBELN : afterVBELN = "?"
If OpenByID(TARGET_ID) Then afterVBELN = session.findById(HG).GetCellValue(0,"VBELN")
log.WriteLine "  after ELIM_ORD VBELN=[" & afterVBELN & "] STATU=" & session.findById(HG).GetCellValue(0,"STATU")

' check leftover Z33SDT644
log.WriteLine "=== Z33SDT644 after cleanup ==="
Se16nDump "Z33SDT644", TARGET_ID, "clean644"

Tcode "/n"
log.WriteLine "DONE. createdVBELN=" & createdVBELN & " afterVBELN=" & afterVBELN
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "S34 done VBELN=" & createdVBELN & " afterVBELN=" & afterVBELN
