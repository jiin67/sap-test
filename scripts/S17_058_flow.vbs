' S17_058_flow.vbs : run ZCRM058*/059*/060* for one dealer, then execute a list of steps on the result screen.
'   args: <tag> <tcode> <kunnr> <spmon YYYY.MM> <vkbur> <NEW|TAB> [step ...]
'   steps (executed in order, all logged):
'     DUMP                 title / status bar / tbar[1] buttons / every ALV grid (columns + up to 80 rows)
'     STATUS               open 시스템 > 상태 popup, dump it (GUI status name), close it
'     SEL:<row>            select row in the first (main) grid
'     BTN:<n>              press wnd[0]/tbar[1]/btn[n]
'     GRIDBTN:<code>       press a toolbar button of the main grid by function code
'     POPDUMP              dump + screenshot wnd[1] (and wnd[2]) without touching them
'     POPBTN:<relid>       press control "wnd[1]/<relid>"  (e.g. usr/btnBUTTON_1)
'     POPSET:<relid>=<v>   set text of "wnd[1]/<relid>"
'     POPKEY:<vkey>        sendVKey on wnd[1]   (0 Enter, 12 Cancel)
'     KEY:<vkey>           sendVKey on wnd[0]   (3 Back)
'     SHOT:<name>          screenshot wnd[0]
'     WAIT:<ms>
'   DB-changing buttons are only pressed when a step says so (user-approved scenario, 2026-09-13).
'   On exit any open popup is cancelled (F12), never confirmed.
'   log: logs\S17_<tag>.txt (UTF-16)   shots: results\shots\S17_<tag>_*.png
'   exit: 0 ok / 11-14 selection screen problem / 15 step failed
'   File must be saved as CP949 (Korean menu texts).
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Dim TAG, TC, KUNNR, SPMON, VKBUR, MODE

If WScript.Arguments.Count < 6 Then
    WScript.Echo "usage: S17_058_flow.vbs <tag> <tcode> <kunnr> <spmon> <vkbur> <NEW|TAB> [step ...]"
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
TC = UCase(WScript.Arguments(1))
KUNNR = WScript.Arguments(2)
SPMON = WScript.Arguments(3)
VKBUR = WScript.Arguments(4)
MODE = UCase(WScript.Arguments(5))

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
Set log = fso.CreateTextFile("C:\sap-test\logs\S17_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG & " tcode=" & TC & " kunnr=" & KUNNR & " spmon=" & SPMON & " vkbur=" & VKBUR & " mode=" & MODE

If connection.Children.Count >= 6 Then
    log.WriteLine "ABORT: session limit reached (" & connection.Children.Count & ")"
    log.Close
    WScript.Quit 14
End If
n0 = connection.Children.Count
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
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" Then WScript.Quit 99
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

Sub Shot(wnd, name)
    On Error Resume Next
    session.findById(wnd).HardCopy "C:\sap-test\results\shots\S17_" & TAG & "_" & name & ".png", 2
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
    WScript.Echo "S17 " & TAG & " exit=" & code
    WScript.Quit code
End Sub

' ---------------- selection screen ----------------
Dim stepNo
stepNo = 0
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & TC
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2500
Err.Clear
On Error GoTo 0
Status "sel:"
If session.Info.ScreenNumber <> 1000 Then Finish 11
Shot "wnd[0]", "0_selection"

On Error Resume Next
If MODE = "TAB" And Exists("wnd[0]/usr/radPA_TAB") Then session.findById("wnd[0]/usr/radPA_TAB").Select
If MODE = "NEW" And Exists("wnd[0]/usr/radPA_NEW") Then session.findById("wnd[0]/usr/radPA_NEW").Select
If Exists("wnd[0]/usr/radPA_DISP") Then session.findById("wnd[0]/usr/radPA_DISP").Select
If Exists("wnd[0]/usr/txtPA_SPMON") Then
    session.findById("wnd[0]/usr/txtPA_SPMON").Text = SPMON
Else
    session.findById("wnd[0]/usr/txtSO_SPMON-LOW").Text = SPMON
End If
If KUNNR <> "*" Then session.findById("wnd[0]/usr/ctxtSO_KUNNR-LOW").Text = KUNNR
If VKBUR <> "*" Then session.findById("wnd[0]/usr/ctxtSO_VKBUR-LOW").Text = VKBUR
If Err.Number <> 0 Then
    log.WriteLine "ERR set selection " & Err.Description
    Err.Clear
    On Error GoTo 0
    Finish 12
End If
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
Err.Clear
On Error GoTo 0
If session.Info.ScreenNumber <> 1000 Then Finish 13
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
Err.Clear
On Error GoTo 0
WScript.Sleep 2500
Status "after F8:"
If session.Info.ScreenNumber = 1000 And Not Exists("wnd[1]") Then
    If session.findById("wnd[0]/sbar").MessageType = "E" Then Finish 14
End If

' ---------------- steps ----------------
Dim a, st, arg, g, eqp, ok
For a = 6 To WScript.Arguments.Count - 1
    stepNo = a - 5
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
        Case "WAIT"
            WScript.Sleep CInt(arg)
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
