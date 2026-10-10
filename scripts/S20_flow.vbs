' S20_flow.vbs : generic step runner (2026-09-14 ZCRMR194 / ZCRM052 / ZCRM053 scenario).
'   Opens a NEW session (SAP Easy Access), runs the steps given as arguments, cancels popups and closes the session.
'   args: <tag> <step> ...   steps: TC: SET:id=v SELECT:id CHK:id=X|- EXISTS:id PRESS:id TREE:n LIST CELL:r,col=v CURCELL:r,col
'         EDITABLE:r GRIDENTER DUMP STATUS SEL:r BTN:n TBTN:n TBTNINFO GRIDBTN:code POPDUMP POPBTN:relid POPSET:relid=v
'         POPKEY:n KEY:n SHOT:name POPSHOT:name WAIT:ms WAITIDLE:s   (id without 'wnd[' prefix is relative to wnd[0]/usr/)
'   DB-changing steps (save etc.) only as the user-approved scenario says. On exit any popup is cancelled (F12).
'   log: logs\S20_<tag>.txt (UTF-16)  shots: results\shots\S20_<tag>_*.png.  File must be saved as CP949.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, hitg, gi, gr, cv
Dim TAG
If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: S20_flow.vbs <tag> <step> ..."
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
' pick the KDQ connection explicitly (KDP/KDD may be open in the same SAP GUI)
Set connection = Nothing
For i = 0 To application.Children.Count - 1
    If application.Children.ElementAt(i).Children.ElementAt(0).Info.SystemName = "KDQ" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then WScript.Quit 99
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder "C:\sap-test\results\shots"
Set log = fso.CreateTextFile("C:\sap-test\logs\S20_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG

If connection.Children.Count >= 6 Then
    log.WriteLine "ABORT: session limit reached (" & connection.Children.Count & ")"
    log.Close
    WScript.Quit 14
End If
n0 = connection.Children.Count
' remember the sessions that exist BEFORE we open ours, then pick the one that is new (2026-10-02: ElementAt(Count-1) picked the user's SE01 session)
Dim oldIds, oi
oldIds = "|"
For oi = 0 To n0 - 1
    oldIds = oldIds & connection.Children.ElementAt(oi).Id & "|"
Next
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
If connection.Children.Count <= n0 Then
    log.WriteLine "ABORT: new session was not created"
    log.Close
    WScript.Quit 14
End If
Set session = Nothing
For oi = 0 To connection.Children.Count - 1
    If InStr(oldIds, "|" & connection.Children.ElementAt(oi).Id & "|") = 0 Then Set session = connection.Children.ElementAt(oi)
Next
If session Is Nothing Then Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" Then WScript.Quit 99
' a freshly created session may still show S000 (logon screen) for a few seconds - wait for the Easy Access menu
For i = 1 To 40
    If session.Info.Transaction = "SESSION_MANAGER" Or session.Info.Transaction = "SMEN" Then Exit For
    WScript.Sleep 500
Next
If session.Info.Transaction <> "SESSION_MANAGER" And session.Info.Transaction <> "SMEN" Then
    log.WriteLine "ABORT: picked session is busy (" & session.Info.Transaction & ") - not ours"
    log.Close
    WScript.Quit 14
End If
log.WriteLine "work session=" & session.Id

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

Sub Status(x)
    On Error Resume Next
    log.WriteLine x & " tcode=" & session.Info.Transaction & " prog=" & session.Info.Program & _
        " screen=" & session.Info.ScreenNumber & " title=[" & T("wnd[0]") & "]" & _
        " sbar=[" & session.findById("wnd[0]/sbar").MessageType & "] " & T("wnd[0]/sbar")
    Err.Clear
    On Error GoTo 0
End Sub

Function Env(n)
    On Error Resume Next
    Env = CreateObject("WScript.Shell").Environment("Process").Item(n)
    Err.Clear
    On Error GoTo 0
End Function

Sub Shot(wnd, name)
    Dim sp, k
    sp = "C:\sap-test\results\shots\S20_" & TAG & "_" & name & ".png"
    On Error Resume Next
    For k = 1 To 3
        session.findById(wnd).HardCopy sp, 2
        If Err.Number <> 0 Then
            log.WriteLine "  SHOT err " & Err.Number & " [" & Err.Description & "] try " & k
            Err.Clear
        End If
        If fso.FileExists(sp) Then Exit For
        WScript.Sleep 600
    Next
    log.WriteLine "  SHOT " & sp & " exists=" & fso.FileExists(sp)
    Err.Clear
    On Error GoTo 0
End Sub

Sub DumpTree(comp, depth, lvl)
    Dim c, k, tp, s
    On Error Resume Next
    tp = comp.Type
    s = String(lvl * 2, " ") & tp & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[") + 1) & " text=[" & comp.Text & "]"
    If tp = "GuiShell" Then s = s & " sub=" & comp.SubType
    log.WriteLine s
    Err.Clear
    If lvl >= depth Then Exit Sub
    k = comp.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    For c = 0 To k - 1
        DumpTree comp.Children.ElementAt(c), depth, lvl + 1
    Next
End Sub

Dim grids()
Dim ng
Sub FindGrids(comp)
    Dim c, k, tp, st
    On Error Resume Next
    tp = comp.Type
    st = ""
    If tp = "GuiShell" Then st = comp.SubType
    Err.Clear
    If st = "GridView" Then
        ReDim Preserve grids(ng)
        Set grids(ng) = comp
        ng = ng + 1
    End If
    k = comp.Children.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    For c = 0 To k - 1
        FindGrids comp.Children.ElementAt(c)
    Next
End Sub

Function MainGrid()
    ng = 0
    FindGrids session.findById("wnd[0]")
    If ng = 0 Then
        Set MainGrid = Nothing
    Else
        Set MainGrid = grids(0)
    End If
End Function

Sub DumpGrid(g, maxRows)
    Dim ord, c, r, line, vis, winTop, nr, hdr
    On Error Resume Next
    log.WriteLine "  GRID " & Mid(g.Id, InStr(g.Id, "/wnd[") + 1) & " rows=" & g.RowCount & " cols=" & g.ColumnCount
    line = ""
    For c = 0 To g.ToolbarButtonCount - 1
        If g.GetToolbarButtonId(c) <> "" Then line = line & " " & g.GetToolbarButtonId(c)
        Err.Clear
    Next
    log.WriteLine "  #TB" & line
    Set ord = g.ColumnOrder
    line = ""
    hdr = ""
    For c = 0 To ord.Count - 1
        If c > 0 Then
            line = line & vbTab
            hdr = hdr & vbTab
        End If
        line = line & ord(c)
        hdr = hdr & g.GetDisplayedColumnTitle(ord(c))
        Err.Clear
    Next
    log.WriteLine "  #COLS" & vbTab & line
    log.WriteLine "  #TITLES" & vbTab & hdr
    nr = g.RowCount
    If nr > maxRows Then nr = maxRows
    vis = g.VisibleRowCount
    If Err.Number <> 0 Or vis < 1 Then
        vis = 1
        Err.Clear
    End If
    winTop = -1
    For r = 0 To nr - 1
        If r < winTop Or r >= winTop + vis Then
            g.FirstVisibleRow = r
            Err.Clear
            winTop = g.FirstVisibleRow
            If Err.Number <> 0 Then
                winTop = r
                Err.Clear
            End If
        End If
        line = ""
        For c = 0 To ord.Count - 1
            If c > 0 Then line = line & vbTab
            line = line & Replace(Replace(g.GetCellValue(r, ord(c)), vbTab, " "), vbCrLf, " ")
            Err.Clear
        Next
        log.WriteLine "  #ROW" & r & vbTab & line
    Next
    Err.Clear
    On Error GoTo 0
End Sub

Sub DoDump()
    Dim k
    Status "  DUMP:"
    For k = 0 To 50
        If Exists("wnd[0]/tbar[1]/btn[" & k & "]") Then
            On Error Resume Next
            log.WriteLine "  btn[" & k & "] text=[" & session.findById("wnd[0]/tbar[1]/btn[" & k & "]").Text & "] tip=[" & _
                          session.findById("wnd[0]/tbar[1]/btn[" & k & "]").Tooltip & "]"
            Err.Clear
            On Error GoTo 0
        End If
    Next
    ng = 0
    FindGrids session.findById("wnd[0]")
    log.WriteLine "  grids=" & ng
    For k = 0 To ng - 1
        DumpGrid grids(k), 80
    Next
    If ng = 0 Then DumpTree session.findById("wnd[0]/usr"), 4, 0
End Sub

Sub DoPopDump()
    Dim w
    For w = 1 To 2
        If Exists("wnd[" & w & "]") Then
            log.WriteLine "  POPUP wnd[" & w & "] title=[" & T("wnd[" & w & "]") & "]"
            DumpTree session.findById("wnd[" & w & "]"), 7, 0
            Shot "wnd[" & w & "]", "popup_w" & w & "_" & stepNo
        End If
    Next
    If Not Exists("wnd[1]") Then log.WriteLine "  (no popup)"
End Sub

' 시스템 > 상태 : find menus by text so that program-specific menus do not matter
Sub DoStatus()
    Dim mb, m, k, j, sub1, found
    found = False
    On Error Resume Next
    Set mb = session.findById("wnd[0]/mbar")
    For k = 0 To mb.Children.Count - 1
        Set m = mb.Children.ElementAt(k)
        If InStr(m.Text, "시스템") > 0 Then
            For j = 0 To m.Children.Count - 1
                Set sub1 = m.Children.ElementAt(j)
                If InStr(sub1.Text, "상태") > 0 Then
                    sub1.Select
                    found = True
                    Exit For
                End If
            Next
        End If
        If found Then Exit For
    Next
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 2000
    If Not found Or Not Exists("wnd[1]") Then
        log.WriteLine "  STATUS: menu or popup not found"
        Exit Sub
    End If
    log.WriteLine "  STATUS popup [" & T("wnd[1]") & "]"
    DumpTree session.findById("wnd[1]"), 8, 0
    Shot "wnd[1]", "status_" & stepNo
    On Error Resume Next
    session.findById("wnd[1]").sendVKey 12
    WScript.Sleep 800
    If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 800
    Err.Clear
    On Error GoTo 0
End Sub

Sub Finish(code)
    Dim k
    On Error Resume Next
    For k = 1 To 3
        If Exists("wnd[1]") Then
            log.WriteLine "  exit: cancel popup [" & T("wnd[1]") & "]"
            session.findById("wnd[1]").sendVKey 12
            WScript.Sleep 800
        End If
    Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    For k = 1 To 3
        If Exists("wnd[1]") Then
            log.WriteLine "  exit: cancel popup [" & T("wnd[1]") & "]"
            session.findById("wnd[1]").sendVKey 12
            WScript.Sleep 800
        End If
    Next
    Err.Clear
    log.WriteLine "exit code=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Echo "S20 " & TAG & " exit=" & code
    WScript.Quit code
End Sub


Function Obj(id)
    If Left(id, 4) = "wnd[" Then
        Set Obj = session.findById(id)
    Else
        Set Obj = session.findById("wnd[0]/usr/" & id)
    End If
End Function

Function IdOf(id)
    If Left(id, 4) = "wnd[" Then
        IdOf = id
    Else
        IdOf = "wnd[0]/usr/" & id
    End If
End Function

Sub DoList()
    Dim c, id, p, q, xy, lines(), k
    ReDim lines(200)
    On Error Resume Next
    For Each c In session.findById("wnd[0]/usr").Children
        id = c.Id
        p = InStrRev(id, "[")
        q = InStrRev(id, "]")
        If p > 0 And q > p Then
            xy = Split(Mid(id, p + 1, q - p - 1), ",")
            If UBound(xy) = 1 Then
                If CInt(xy(1)) <= 200 And Trim(c.Text) <> "" Then lines(CInt(xy(1))) = lines(CInt(xy(1))) & " [" & xy(0) & "]" & c.Text
            End If
        End If
    Next
    Err.Clear
    On Error GoTo 0
    For k = 0 To 200
        If lines(k) <> "" Then log.WriteLine "  y" & k & ":" & lines(k)
    Next
End Sub

Dim stepNo, a, st, arg, g, eqp, ok, o, rc, ordc, cc, ed, wi
stepNo = 0
For a = 1 To WScript.Arguments.Count - 1
    stepNo = a
    st = WScript.Arguments(a)
    arg = ""
    If InStr(st, ":") > 0 Then
        arg = Mid(st, InStr(st, ":") + 1)
        st = Left(st, InStr(st, ":") - 1)
    End If
    log.WriteLine "==== STEP " & stepNo & " " & UCase(st) & " " & arg
    ok = True
    On Error Resume Next
    Select Case UCase(st)
        Case "TC"
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & arg
            session.findById("wnd[0]").sendVKey 0
            WScript.Sleep 3000
            If UCase(Trim(Env("S20_NOMAX"))) <> "1" Then
                session.findById("wnd[0]").maximize
                WScript.Sleep 500
            End If
        Case "SET"
            eqp = InStr(arg, "=")
            Obj(Left(arg, eqp - 1)).Text = Mid(arg, eqp + 1)
            log.WriteLine "  set " & Left(arg, eqp - 1) & " = " & Mid(arg, eqp + 1)
        Case "SELECT"
            Obj(arg).Select
            WScript.Sleep 800
        Case "COMBO"
            eqp = InStr(arg, "=")
            Obj(Left(arg, eqp - 1)).Key = Mid(arg, eqp + 1)
            log.WriteLine "  combo " & Left(arg, eqp - 1) & " key = " & Mid(arg, eqp + 1)
            WScript.Sleep 800
        Case "CHK"
            eqp = InStr(arg, "=")
            Obj(Left(arg, eqp - 1)).Selected = (Mid(arg, eqp + 1) = "X")
        Case "EXISTS"
            log.WriteLine "  EXISTS " & arg & " = " & Exists(IdOf(arg))
        Case "PRESS"
            Obj(arg).press
            WScript.Sleep 3000
        Case "TREE"
            On Error GoTo 0
            DumpTree session.findById("wnd[0]"), CInt(arg), 0
        Case "MENU"
            log.WriteLine "  menu " & arg & " [" & session.findById(arg).Text & "]"
            session.findById(arg).select
            WScript.Sleep 2000
        Case "FOCUS"
            Obj(arg).SetFocus
            log.WriteLine "  focus " & arg & " [" & Obj(arg).Text & "]"
            WScript.Sleep 300
        Case "LIST"
            On Error GoTo 0
            DoList
        Case "CELL"
            eqp = InStr(arg, "=")
            rc = Split(Left(arg, eqp - 1), ",")
            Set g = MainGrid()
            g.modifyCell CLng(rc(0)), rc(1), Mid(arg, eqp + 1)
            log.WriteLine "  modifyCell " & rc(0) & "," & rc(1) & " = " & Mid(arg, eqp + 1)
        Case "SELCOL"
            ' SELCOL:<columnId>  -> select a whole ALV column (needed before &SORT_ASC / &SORT_DSC / &MB_FILTER)
            Set g = MainGrid()
            g.selectedColumns = arg
            WScript.Sleep 500
            log.WriteLine "  selectedColumns set to " & arg
        Case "CHKBOX"
            ' CHKBOX:<row>,<col>=1|0  -> tick/untick an ALV checkbox cell (modifyCheckbox)
            eqp = InStr(arg, "=")
            rc = Split(Left(arg, eqp - 1), ",")
            Set g = MainGrid()
            g.setCurrentCell CLng(rc(0)), rc(1)
            g.modifyCheckbox CLng(rc(0)), rc(1), (Mid(arg, eqp + 1) = "1")
            WScript.Sleep 500
            log.WriteLine "  modifyCheckbox " & rc(0) & "," & rc(1) & " = " & Mid(arg, eqp + 1) & " now=[" & g.GetCellValue(CLng(rc(0)), rc(1)) & "]"
        Case "CELLALL"
            ' CELLALL:<col>=<value>  -> every row of the main grid
            eqp = InStr(arg, "=")
            Set g = MainGrid()
            For cc = 0 To g.RowCount - 1
                g.modifyCell cc, Left(arg, eqp - 1), Mid(arg, eqp + 1)
            Next
            log.WriteLine "  modifyCell all " & g.RowCount & " rows " & arg
        Case "EXEC"
            ' EXEC:<command line>  -> run synchronously while this session stays on its screen
            rc = CreateObject("WScript.Shell").Run(arg, 0, True)
            log.WriteLine "  exec rc=" & rc
            If rc <> 0 Then ok = False
        Case "SELROWS"
            ' SELROWS:<col>=<v1>|<v2>...  -> select the rows whose column value contains one of the values
            eqp = InStr(arg, "=")
            Set g = MainGrid()
            ed = ""
            For cc = 0 To g.RowCount - 1
                For Each o In Split(Mid(arg, eqp + 1), "|")
                    If InStr(g.GetCellValue(cc, Left(arg, eqp - 1)), o) > 0 Then
                        If ed <> "" Then ed = ed & ","
                        ed = ed & cc
                    End If
                Next
            Next
            log.WriteLine "  selected rows=" & ed
            If ed = "" Then
                ok = False
            Else
                g.selectedRows = ed
                WScript.Sleep 500
            End If
        Case "CURCELL"
            rc = Split(arg, ",")
            Set g = MainGrid()
            g.currentCellRow = CLng(rc(0))
            g.currentCellColumn = rc(1)
        Case "EDITABLE"
            Set g = MainGrid()
            Set ordc = g.ColumnOrder
            ed = ""
            For cc = 0 To ordc.Count - 1
                If g.GetCellChangeable(CLng(arg), ordc(cc)) Then ed = ed & " " & ordc(cc)
                Err.Clear
            Next
            log.WriteLine "  changeable cols row " & arg & ":" & ed
        Case "GRIDENTER"
            Set g = MainGrid()
            g.pressEnter
            WScript.Sleep 1500
        Case "DUMP"
            On Error GoTo 0
            DoDump
        Case "STATUS"
            On Error GoTo 0
            DoStatus
        Case "SEL"
            Set g = MainGrid()
            If g Is Nothing Then
                ok = False
            Else
                g.selectedRows = arg
                WScript.Sleep 500
            End If
        Case "BTN"
            If Exists("wnd[0]/tbar[1]/btn[" & arg & "]") Then
                log.WriteLine "  press btn[" & arg & "] [" & session.findById("wnd[0]/tbar[1]/btn[" & arg & "]").Text & "]"
                session.findById("wnd[0]/tbar[1]/btn[" & arg & "]").press
                WScript.Sleep 3000
            Else
                log.WriteLine "  btn[" & arg & "] not present"
                ok = False
            End If
        Case "TBTN"
            If Exists("wnd[0]/tbar[0]/btn[" & arg & "]") Then
                log.WriteLine "  tbar0 btn[" & arg & "] changeable=" & session.findById("wnd[0]/tbar[0]/btn[" & arg & "]").Changeable
                session.findById("wnd[0]/tbar[0]/btn[" & arg & "]").press
                WScript.Sleep 3000
            Else
                log.WriteLine "  tbar0 btn[" & arg & "] not present"
                ok = False
            End If
        Case "TBTNINFO"
            For cc = 0 To 50
                If Exists("wnd[0]/tbar[0]/btn[" & cc & "]") Then log.WriteLine "  tbar0 btn[" & cc & "] changeable=" & session.findById("wnd[0]/tbar[0]/btn[" & cc & "]").Changeable & " tip=[" & session.findById("wnd[0]/tbar[0]/btn[" & cc & "]").Tooltip & "]"
                Err.Clear
            Next
        Case "SCROLLCOL"
            Set g = MainGrid()
            g.setCurrentCell 0, arg
            WScript.Sleep 800
        Case "GRIDMENU"
            rc = Split(arg, ",")
            Set g = MainGrid()
            g.pressToolbarContextButton rc(0)
            WScript.Sleep 800
            g.selectContextMenuItem rc(1)
            WScript.Sleep 2500
        Case "CALDATE"
            ' CALDATE:<yyyymmdd>  -> pick a date in the F4_DATE calendar popup (wnd[1])
            Set o = session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell")
            log.WriteLine "  calendar type=" & o.Type & " sub=" & o.SubType & " err=" & Err.Number
            Err.Clear
            o.FocusDate = arg
            log.WriteLine "  set FocusDate err=" & Err.Number & " " & Err.Description
            Err.Clear
            o.SelectionInterval = arg & "," & arg
            log.WriteLine "  set SelectionInterval err=" & Err.Number & " " & Err.Description
            Err.Clear
            WScript.Sleep 800
            log.WriteLine "  calendar focus=" & o.FocusDate & " sel=" & o.SelectionInterval & " err=" & Err.Number
            Err.Clear
        Case "POPSELCOL"
            ' POPSELCOL:<column title>  -> find that row in ANY popup grid (SELTEXT) and select it
            ng = 0
            FindGrids session.findById("wnd[1]")
            hitg = -1
            For gi = 0 To ng - 1
                If hitg < 0 Then
                    For gr = 0 To grids(gi).RowCount - 1
                        On Error Resume Next
                        cv = grids(gi).GetCellValue(gr, "SELTEXT")
                        If Err.Number <> 0 Then
                            Err.Clear
                            cv = ""
                        End If
                        On Error GoTo 0
                        If Trim(cv) = Trim(arg) And hitg < 0 Then
                            hitg = gi
                            grids(gi).setCurrentCell CLng(gr), "SELTEXT"
                            grids(gi).selectedRows = CStr(gr)
                            log.WriteLine "  POPSELCOL [" & arg & "] grid=" & gi & " row=" & gr & " sel=[" & grids(gi).selectedRows & "]"
                        End If
                    Next
                End If
            Next
            If hitg < 0 Then
                log.WriteLine "  POPSELCOL [" & arg & "] NOT FOUND in " & ng & " grids"
                ok = False
            End If
            WScript.Sleep 600
        Case "POPSELROW"
            ' POPSELROW:<gridIndex>,<row>  -> select one row in a popup grid (layout dialog etc.)
            rc = Split(arg, ",")
            ng = 0
            FindGrids session.findById("wnd[1]")
            log.WriteLine "  popup grids=" & ng & " pick=" & rc(0)
            grids(CInt(rc(0))).selectedRows = rc(1)
            grids(CInt(rc(0))).setCurrentCell CLng(rc(1)), grids(CInt(rc(0))).ColumnOrder(0)
            WScript.Sleep 800
            log.WriteLine "  selected rows=[" & grids(CInt(rc(0))).selectedRows & "]"
        Case "POPSCROLL"
            ng = 0
            FindGrids session.findById("wnd[1]")
            grids(0).FirstVisibleRow = CLng(arg)
            WScript.Sleep 800
        Case "POPTREE"
            On Error GoTo 0
            DumpTree session.findById("wnd[1]"), CInt(arg), 0
        Case "POPGRIDS"
            ng = 0
            FindGrids session.findById("wnd[1]")
            log.WriteLine "  popup grids=" & ng
            For cc = 0 To ng - 1
                DumpGrid grids(cc), 120
            Next
        Case "GRIDBTN"
            Set g = MainGrid()
            g.pressToolbarButton arg
            WScript.Sleep 2500
        Case "POPDUMP"
            On Error GoTo 0
            DoPopDump
        Case "POPBTN"
            If Exists("wnd[1]/" & arg) Then
                log.WriteLine "  press wnd[1]/" & arg & " [" & T("wnd[1]/" & arg) & "]"
                session.findById("wnd[1]/" & arg).press
                WScript.Sleep 3000
            Else
                log.WriteLine "  wnd[1]/" & arg & " not present"
                ok = False
            End If
        Case "POPSET"
            eqp = InStr(arg, "=")
            session.findById("wnd[1]/" & Left(arg, eqp - 1)).Text = Mid(arg, eqp + 1)
            log.WriteLine "  set wnd[1]/" & Left(arg, eqp - 1) & " = " & Mid(arg, eqp + 1)
        Case "POPKEY"
            If Exists("wnd[1]") Then
                session.findById("wnd[1]").sendVKey CInt(arg)
                WScript.Sleep 2500
            Else
                log.WriteLine "  no popup for POPKEY"
                ok = False
            End If
        Case "KEY"
            session.findById("wnd[0]").sendVKey CInt(arg)
            WScript.Sleep 2500
        Case "SHOT"
            On Error GoTo 0
            Shot "wnd[0]", arg
        Case "POPSHOT"
            On Error GoTo 0
            Shot "wnd[1]", arg
        Case "WAIT"
            WScript.Sleep CLng(arg)
        Case "WAITIDLE"
            For wi = 1 To CLng(arg)
                If Not session.Busy Then Exit For
                WScript.Sleep 1000
            Next
            log.WriteLine "  waited " & wi & "s busy=" & session.Busy
        Case Else
            log.WriteLine "  unknown step"
            ok = False
    End Select
    If Err.Number <> 0 Then
        log.WriteLine "  ERR " & Err.Number & " " & Err.Description
        Err.Clear
        ok = False
    End If
    On Error GoTo 0
    Status "  after step " & stepNo & ":"
    If Exists("wnd[1]") Then log.WriteLine "  (popup open: [" & T("wnd[1]") & "])"
    If Not ok Then
        log.WriteLine "  STEP FAILED - stop"
        Finish 15
    End If
Next
Finish 0

