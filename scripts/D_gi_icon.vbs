' D_gi_icon.vbs : discovery for scenario 9-1 (GI_ICON column). READ ONLY.
'  1) header ALV columns/titles -> GI_ICON present? title? ; toolbar ids (new GI button?)
'  2) wide query -> paged read of BSTKD/VBELN/VBELN_VL/GI_ICON ; tally icon values for rows with delivery
'  3) SE16N VBUP WBSTA for sample deliveries per icon state
'  4) layout: change-layout popup (hidden columns) + saved layouts list ; screenshots
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder("C:\sap-test\results\shots")
Set log = fso.CreateTextFile("C:\sap-test\logs\D_gi_icon.txt", True, True)
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
Const SHOTS = "C:\sap-test\results\shots\"
Dim K_GI : K_GI = ChrW(&HCD9C)&ChrW(&HACE0)&ChrW(&HC804)&ChrW(&HAE30)   ' chulgo-jeongi
Dim K_DOC : K_DOC = ChrW(&HBB38)&ChrW(&HC11C)                           ' munseo (document)

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
Function WT(n)
    On Error Resume Next
    WT = "" : WT = session.findById("wnd[" & n & "]").Text
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
Sub ShotPopup(name)
    On Error Resume Next
    session.findById("wnd[1]").HardCopy SHOTS & name, 2
    Err.Clear
    On Error GoTo 0
End Sub
Sub Dump(comp, depth)
    Dim c, k, line
    line = String(depth*2," ") & comp.Type
    On Error Resume Next
    line = line & "/" & comp.SubType & " | " & comp.Id & " | [" & Left(comp.Text,70) & "]"
    On Error GoTo 0
    log.WriteLine line
    On Error Resume Next
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k-1 : Dump comp.Children.ElementAt(c), depth+1 : Next
End Sub
Function OpenQuery(bstkd, fromD, toD)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = bstkd
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = fromD
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = toD
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 6000
    OpenQuery = (session.Info.ScreenNumber = 100)
End Function

' ---------- 1) columns + toolbar ----------
If Not OpenQuery("", "2026.01.01", "2026.09.16") Then log.WriteLine "no result screen: " & Sbar() : log.Close : WScript.Quit 10
Dim g : Set g = session.findById(HG)
Dim cols : Set cols = g.ColumnOrder
Dim c, hasGI, giPos, giTitle
hasGI = False : giPos = -1 : giTitle = ""
log.WriteLine "=== header ALV columns (" & cols.Count & ") rows=" & g.RowCount & " ==="
For c = 0 To cols.Count - 1
    Dim t : t = ""
    On Error Resume Next
    t = g.GetColumnTitles(cols(c))(0)
    On Error GoTo 0
    log.WriteLine "  col" & c & " " & cols(c) & " | " & t
    If cols(c) = "GI_ICON" Then hasGI = True : giPos = c : giTitle = t
Next
log.WriteLine "GI_ICON present=" & hasGI & " pos=" & giPos & " title=[" & giTitle & "] titleIs'chulgo-jeongi'=" & (giTitle = K_GI)
log.WriteLine "=== toolbar ==="
Dim b
For b = 0 To g.ToolbarButtonCount - 1
    log.WriteLine "  tb" & b & " id=[" & g.GetToolbarButtonId(b) & "] type=" & g.GetToolbarButtonType(b)
Next
' screenshot with GI_ICON column scrolled into view
On Error Resume Next
If hasGI Then g.FirstVisibleColumn = "VBELN"
Err.Clear
On Error GoTo 0
WScript.Sleep 500
Shot "ALV_01_header_gi_column.png"

' ---------- 2) paged read ----------
Dim n : n = g.RowCount
Dim vis : vis = g.VisibleRowCount
If vis < 1 Then vis = 16
Dim withVL, iconTally : withVL = 0
Set iconTally = CreateObject("Scripting.Dictionary")
Dim samples : Set samples = CreateObject("Scripting.Dictionary")   ' icon -> "bstkd|vbeln|vl;..."
Dim page, r, bs, vb, vl, ic, key
For page = 0 To n - 1 Step vis
    Set g = session.findById(HG)
    g.FirstVisibleRow = page
    WScript.Sleep 200
    Set g = session.findById(HG)
    For r = page To page + vis - 1
        If r >= n Then Exit For
        vl = g.GetCellValue(r, "VBELN_VL")
        If vl <> "" Then
            withVL = withVL + 1
            bs = g.GetCellValue(r, "BSTKD") : vb = g.GetCellValue(r, "VBELN")
            ic = "" : If hasGI Then ic = g.GetCellValue(r, "GI_ICON")
            key = "[" & ic & "]"
            If iconTally.Exists(key) Then iconTally(key) = iconTally(key) + 1 Else iconTally.Add key, 1
            If Not samples.Exists(key) Then samples.Add key, ""
            If Len(samples(key)) < 400 Then samples(key) = samples(key) & bs & "|" & vb & "|" & vl & "|" & g.GetCellValue(r,"ID") & ";"
        End If
    Next
Next
log.WriteLine "=== rows with delivery: " & withVL & " of " & n & " ==="
Dim k
For Each k In iconTally.Keys
    log.WriteLine "  icon " & k & " count=" & iconTally(k)
    log.WriteLine "     samples: " & samples(k)
Next

' ---------- 3) VBUP WBSTA for up to 2 samples per icon state ----------
Sub VbupCheck(vbelnVL, label)
    Tcode "/nSE16N"
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "VBUP"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    If session.findById("wnd[0]/usr/chkGD-EDIT").Selected Then session.findById("wnd[0]/usr/chkGD-EDIT").Selected = False
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "200"
    Dim rr, desc, setC : setC = False
    For rr = 0 To 5
        desc = "" : desc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & rr & "]").Text
        If Err.Number <> 0 Then Err.Clear : Exit For
        If InStr(desc, K_DOC) > 0 Then
            session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2," & rr & "]").Text = vbelnVL
            setC = True : log.WriteLine "  (VBUP cond row " & rr & " desc=[" & desc & "])" : Exit For
        End If
    Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    WScript.Sleep 2200
    Dim gg, r2, ln
    Set gg = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    If Err.Number <> 0 Then log.WriteLine "  [" & label & "] VBUP " & vbelnVL & " condSet=" & setC & " -> no ALV, sbar=" & Sbar() : Err.Clear : On Error GoTo 0 : Exit Sub
    ln = "  [" & label & "] VBUP " & vbelnVL & " rows=" & gg.RowCount & " :"
    For r2 = 0 To gg.RowCount - 1
        If r2 > 12 Then Exit For
        ln = ln & " " & gg.GetCellValue(r2,"POSNR") & "=" & gg.GetCellValue(r2,"WBSTA")
    Next
    log.WriteLine ln
    On Error GoTo 0
End Sub
log.WriteLine "=== VBUP checks ==="
For Each k In samples.Keys
    Dim parts, p, cnt
    parts = Split(samples(k), ";")
    cnt = 0
    For p = 0 To UBound(parts)
        If parts(p) <> "" And cnt < 2 Then
            Dim f : f = Split(parts(p), "|")
            VbupCheck f(2), "icon" & k & " bstkd=" & f(0) & " vbeln=" & f(1)
            cnt = cnt + 1
        End If
    Next
Next

' ---------- 4) layout popups ----------
log.WriteLine "=== layout ==="
If OpenQuery("", "2026.09.01", "2026.09.16") Then
    Set g = session.findById(HG)
    On Error Resume Next
    g.pressToolbarContextButton "&MB_VARIANT"
    WScript.Sleep 600
    g.selectContextMenuItem "&LOAD"
    WScript.Sleep 1500
    If HasW(1) Then
        log.WriteLine "  saved-layout popup=[" & WT(1) & "]"
        Dump session.findById("wnd[1]/usr"), 2
        ShotPopup "ALV_08_saved_layouts.png"
    Else
        log.WriteLine "  &LOAD: no popup (no saved layouts?) sbar=" & Sbar()
    End If
    CloseAll
    Set g = session.findById(HG)
    g.pressToolbarContextButton "&MB_VARIANT"
    WScript.Sleep 600
    g.selectContextMenuItem "&COL0"
    WScript.Sleep 1500
    If HasW(1) Then
        log.WriteLine "  change-layout popup=[" & WT(1) & "]"
        Dump session.findById("wnd[1]/usr"), 2
        ShotPopup "ALV_08_change_layout.png"
    Else
        log.WriteLine "  &COL0: no popup sbar=" & Sbar()
    End If
    CloseAll
    Err.Clear
    On Error GoTo 0
End If

Tcode "/n"
log.WriteLine "done"
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "D_gi_icon done GI=" & hasGI & " withVL=" & withVL
