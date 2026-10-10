' D_gi_kdq.vbs : initial-state recording on KDQ for the 2026-09-16 scenario (READ ONLY).
'  - attaches to the connection whose SystemName = KDQ (several connections may be open)
'  - selection screen dump (DEL_TEST directory field?) ; header ALV columns + toolbar (ALV-01)
'  - per target BSTKD: headers + items, screenshot (BAS-01/02)
'  - VBUP WBSTA for existing deliveries ; MARA MEINS + MVKE KTGRM per material (BAS-03)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i, ci
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For ci = 0 To application.Children.Count - 1
    If application.Children.ElementAt(ci).Children.ElementAt(0).Info.SystemName = "KDQ" Then Set connection = application.Children.ElementAt(ci) : Exit For
Next
If connection Is Nothing Then WScript.Echo "NO KDQ CONNECTION" : WScript.Quit 99
Set base = connection.Children.ElementAt(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots\gi") Then fso.CreateFolder("C:\sap-test\results\shots\gi")
Set log = fso.CreateTextFile("C:\sap-test\logs\D_gi_kdq.txt", True, True)
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
If session.Info.SystemName <> "KDQ" Then log.WriteLine "ABORT not KDQ" : log.Close : WScript.Quit 99
Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"
Const SHOTS = "C:\sap-test\results\shots\gi\"
Dim K_MATNR : K_MATNR = ChrW(&HC790)&ChrW(&HC7AC)&ChrW(&HBC88)&ChrW(&HD638)
Dim K_DOC : K_DOC = ChrW(&HBB38)&ChrW(&HC11C)
Dim TARGETS : TARGETS = Array("1757", "2318", "1722", "2392", "2766", "2827")

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Function Sbar() : Sbar = session.findById("wnd[0]/sbar").Text : End Function
Function HasW(n)
    On Error Resume Next
    HasW = False
    Dim w : Set w = session.findById("wnd[" & n & "]")
    If Err.Number = 0 Then HasW = True
    Err.Clear
    On Error GoTo 0
End Function
Sub CloseAll()
    Dim g, n
    On Error Resume Next
    For g = 1 To 6
        If Not HasW(1) Then Exit Sub
        For n = 3 To 1 Step -1
            If HasW(n) Then session.findById("wnd[" & n & "]/tbar[0]/btn[12]").press : Err.Clear
            If HasW(n) Then session.findById("wnd[" & n & "]").close : Err.Clear
        Next
        WScript.Sleep 400
    Next
    On Error GoTo 0
End Sub
Sub Shot(name)
    On Error Resume Next
    session.findById("wnd[0]").HardCopy SHOTS & name, 2
    Err.Clear
    On Error GoTo 0
End Sub
Sub DumpUsr()
    Dim usr, c, comp, line
    Set usr = session.findById("wnd[0]/usr")
    For c = 0 To usr.Children.Count - 1
        Set comp = usr.Children.ElementAt(c)
        If comp.Type <> "GuiLabel" Or InStr(comp.Name, "%_") = 0 Then
            line = "  " & comp.Type & " | " & comp.Name
            On Error Resume Next
            line = line & " | [" & comp.Text & "]"
            On Error GoTo 0
            log.WriteLine line
        End If
    Next
End Sub
Function OpenBSTKD(b)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = b
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 3500
    OpenBSTKD = (session.Info.ScreenNumber = 100)
End Function
Sub Se16n(table, condDesc, val, colsCsv, label, maxRows)
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "500"
    Dim rr, desc, setC, used : setC = False : used = ""
    For rr = 0 To 8
        desc = "" : desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, condDesc) > 0 Then session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = val : setC = True : used = desc : Exit For
    Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2000
    Dim g, r2, ln, cc, cs
    cs = Split(colsCsv, ",")
    Set g = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then log.WriteLine "  [" & label & "] " & table & "=" & val & " cond=" & setC & "(" & used & ") -> no ALV: " & Sbar() : Err.Clear : On Error GoTo 0 : Exit Sub
    ln = "  [" & label & "] " & table & "=" & val & " rows=" & g.RowCount & " :"
    For r2 = 0 To g.RowCount - 1
        If r2 >= maxRows Then Exit For
        ln = ln & " {"
        For cc = 0 To UBound(cs)
            ln = ln & cs(cc) & "=" & g.GetCellValue(r2, cs(cc)) & " "
        Next
        ln = ln & "}"
    Next
    log.WriteLine ln
    On Error GoTo 0
End Sub

' ---------- selection screen ----------
Tcode "/nZSD600"
log.WriteLine "=== selection screen (P_RBLIST) ==="
DumpUsr
Shot "SEL_screen.png"

' ---------- columns / toolbar (ALV-01) ----------
Dim mats : Set mats = CreateObject("Scripting.Dictionary")
Dim dels : Set dels = CreateObject("Scripting.Dictionary")
Dim t, gh, gi, cols, c, hasGI
If OpenBSTKD(TARGETS(0)) Then
    Set gh = session.findById(HG)
    Set cols = gh.ColumnOrder
    hasGI = False
    log.WriteLine "=== header columns (" & cols.Count & ") ==="
    For c = 0 To cols.Count - 1
        Dim tt : tt = "" : On Error Resume Next : tt = gh.GetColumnTitles(cols(c))(0) : On Error GoTo 0
        log.WriteLine "  " & c & " " & cols(c) & " | " & tt
        If cols(c) = "GI_ICON" Then hasGI = True
    Next
    log.WriteLine "GI_ICON present=" & hasGI
    Dim b
    For b = 0 To gh.ToolbarButtonCount - 1
        If gh.GetToolbarButtonId(b) <> "" Then log.WriteLine "  tb" & b & " " & gh.GetToolbarButtonId(b)
    Next
End If

' ---------- per BSTKD state ----------
Dim k
For Each t In TARGETS
    log.WriteLine "=== BSTKD " & t & " ==="
    If Not OpenBSTKD(t) Then
        log.WriteLine "  no result: " & Sbar()
    Else
        Set gh = session.findById(HG)
        Dim h
        For h = 0 To gh.RowCount - 1
            Dim gicon : gicon = "" : If hasGI Then gicon = gh.GetCellValue(h,"GI_ICON")
            log.WriteLine "  H" & h & ": ID=" & gh.GetCellValue(h,"ID") & " VER=" & gh.GetCellValue(h,"VERSION") & " STATU=" & gh.GetCellValue(h,"STATU") & " VBELN=[" & gh.GetCellValue(h,"VBELN") & "] VBELN_VL=[" & gh.GetCellValue(h,"VBELN_VL") & "] GI_ICON=[" & gicon & "] VDATU=" & gh.GetCellValue(h,"VDATU") & " KUNNR=" & gh.GetCellValue(h,"KUNNR") & " AUART=" & gh.GetCellValue(h,"AUART") & " LOEVM=[" & gh.GetCellValue(h,"LOEVM") & "]"
            If gh.GetCellValue(h,"VBELN_VL") <> "" Then dels(gh.GetCellValue(h,"VBELN_VL")) = t
            gh.setCurrentCell h,"BSTKD" : gh.selectedRows = CStr(h)
            On Error Resume Next
            gh.doubleClickCurrentCell
            Err.Clear
            On Error GoTo 0
            WScript.Sleep 1600
            CloseAll
            log.WriteLine "     dblclick sbar=[" & Sbar() & "]"
            Set gi = session.findById(IG)
            Dim r
            For r = 0 To gi.RowCount - 1
                log.WriteLine "     I" & r & ": SEQNO=" & gi.GetCellValue(r,"SEQNO") & " MATNR=" & gi.GetCellValue(r,"MATNR") & " MAKTX=[" & gi.GetCellValue(r,"MAKTX") & "] ZMENG=" & gi.GetCellValue(r,"ZMENG") & " MEINS=" & gi.GetCellValue(r,"MEINS") & " LGORT=" & gi.GetCellValue(r,"LGORT") & " KDKG1=" & gi.GetCellValue(r,"KDKG1")
                If gi.GetCellValue(r,"MATNR") <> "" Then mats(gi.GetCellValue(r,"MATNR")) = 1
            Next
            If h = 0 Then Shot "BAS_" & t & "_items.png"
        Next
    End If
Next

' ---------- VBUP for existing deliveries ----------
log.WriteLine "=== VBUP (existing deliveries) ==="
For Each k In dels.Keys
    Se16n "VBUP", K_DOC, k, "VBELN,POSNR,WBSTA,LFSTA,FKSTA", "VL " & k & " (BSTKD " & dels(k) & ")", 12
Next

' ---------- MARA / MVKE per material (BAS-03) ----------
log.WriteLine "=== MARA / MVKE ==="
For Each k In mats.Keys
    Se16n "MARA", K_MATNR, k, "MATNR,MTART,MEINS,LVORM", "MARA", 1
    Se16n "MVKE", K_MATNR, k, "MATNR,VKORG,VTWEG,KTGRM,LVORM", "MVKE", 6
Next

Tcode "/n"
log.WriteLine "done"
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "D_gi_kdq done materials=" & mats.Count & " deliveries=" & dels.Count
