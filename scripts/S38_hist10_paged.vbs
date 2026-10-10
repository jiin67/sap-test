' S38_hist10_paged.vbs : HIST-10 with proper ALV paging (FirstVisibleRow) - read ALL header rows. Read only.
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Echo "GUARD" : WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S38_hist10_paged.txt", True, True)
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
Const MAT = "GXXX002868"

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
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
            If HasW(n) Then session.findById("wnd[" & n & "]").sendVKey 0 : WScript.Sleep 300 : If HasW(n) Then session.findById("wnd[" & n & "]").close
            Err.Clear
        Next
        WScript.Sleep 300
    Next
    On Error GoTo 0
End Sub
Function OpenQuery(bstkd, matnr)
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_ID-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = bstkd
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = matnr
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 4000
    OpenQuery = (session.Info.ScreenNumber = 100)
End Function
' read a column for ALL rows with paging
Function ReadAll(gridId, col, n)
    Dim g, vis, r, arr, page
    ReDim arr(n - 1)
    Set g = session.findById(gridId)
    vis = g.VisibleRowCount
    If vis < 1 Then vis = 10
    For page = 0 To n - 1 Step vis
        Set g = session.findById(gridId)
        g.FirstVisibleRow = page
        WScript.Sleep 250
        Set g = session.findById(gridId)
        For r = page To page + vis - 1
            If r >= n Then Exit For
            arr(r) = g.GetCellValue(r, col)
        Next
    Next
    ReadAll = arr
End Function

If Not OpenQuery("", MAT) Then log.WriteLine "no result" : log.Close : WScript.Quit 10
Dim gh : Set gh = session.findById(HG)
Dim n : n = gh.RowCount
log.WriteLine "material query headerRows=" & n & " visibleRowCount=" & gh.VisibleRowCount
Dim ids, bs, vs, sts
ids = ReadAll(HG, "ID", n)
bs  = ReadAll(HG, "BSTKD", n)
vs  = ReadAll(HG, "VERSION", n)
sts = ReadAll(HG, "STATU", n)
Dim d : Set d = CreateObject("Scripting.Dictionary")
Dim r, key, dup, empties : dup = 0 : empties = 0
For r = 0 To n - 1
    key = ids(r) & "|" & bs(r) & "|" & vs(r)
    If ids(r) = "" Then empties = empties + 1
    If d.Exists(key) Then d(key) = d(key) & "," & r : dup = dup + 1 Else d.Add key, CStr(r)
Next
log.WriteLine "uniqueKeys=" & d.Count & " duplicateRows=" & dup & " emptyIdRows=" & empties
Dim k
For Each k In d.Keys
    If InStr(d(k), ",") > 0 Then log.WriteLine "  DUP " & k & " rows " & d(k)
Next
For r = 0 To n - 1
    log.WriteLine "  " & r & ": BSTKD=" & bs(r) & " VER=" & vs(r) & " STATU=" & sts(r) & " ID=" & ids(r)
Next
If dup = 0 And empties = 0 Then log.WriteLine "[PASS] HIST-10 all " & n & " header rows unique by ID|BSTKD|VERSION" Else log.WriteLine "[FAIL] HIST-10 dup=" & dup & " empties=" & empties
' cross-check: a header with >1 item of this material would prove de-dup is working
Dim best, bestCnt, h : best = "" : bestCnt = 0
For h = 0 To n - 1
    If h > 7 Then Exit For
    Set gh = session.findById(HG)
    gh.FirstVisibleRow = 0
    gh.setCurrentCell h, "BSTKD" : gh.selectedRows = CStr(h)
    On Error Resume Next
    gh.doubleClickCurrentCell
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 1500
    CloseAll
    Dim gi : Set gi = session.findById(IG)
    Dim m : Set m = Nothing
    Dim mats : mats = ReadAll(IG, "MATNR", gi.RowCount)
    Dim c, cnt : cnt = 0
    For c = 0 To gi.RowCount - 1
        If mats(c) = MAT Then cnt = cnt + 1
    Next
    log.WriteLine "  header " & h & " (" & bs(h) & "): items=" & gi.RowCount & " itemsWith" & MAT & "=" & cnt
    If cnt > bestCnt Then bestCnt = cnt : best = bs(h)
Next
log.WriteLine "max items with material in one header=" & bestCnt & " (BSTKD " & best & ")"
Tcode "/n"
log.Close
If Not HasW(1) Then session.findById("wnd[0]").close
WScript.Echo "S38 done rows=" & n & " dup=" & dup & " empties=" & empties
