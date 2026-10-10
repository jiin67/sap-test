' inspect_058_result.vbs : READ-ONLY. Run ZCRM058*/059*/060* for one dealer and dump the result screen.
'   args: <tcode> <kunnr> <spmon YYYY.MM> <tag> [<vkbur>]
'   Presses only F8 on the selection screen. Popups are logged and CANCELLED (F12), never confirmed.
'   log : logs\inspect_058_<tag>.txt (UTF-16), shot : results\shots\SUBS_<tag>.png
'   exit: 0 ok / 11-13 selection screen problem / 14 F8 rejected with an E message
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Dim TC, KUNNR, SPMON, TAG, VKBUR

If WScript.Arguments.Count < 4 Then
    WScript.Echo "usage: inspect_058_result.vbs <tcode> <kunnr> <spmon> <tag> [<vkbur>]"
    WScript.Quit 1
End If
TC = UCase(WScript.Arguments(0))
KUNNR = WScript.Arguments(1)
SPMON = WScript.Arguments(2)
TAG = WScript.Arguments(3)
VKBUR = ""
If WScript.Arguments.Count >= 5 Then VKBUR = WScript.Arguments(4)

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
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_058_" & TAG & ".txt", True, True)
log.WriteLine "tcode=" & TC & " kunnr=" & KUNNR & " spmon=" & SPMON

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" And session.Info.SystemName <> "KDD" Then WScript.Quit 99
log.WriteLine "work session=" & session.Id

Function Exists(id)
    Dim o
    On Error Resume Next
    Set o = session.findById(id)
    Exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Sub Status(tagx)
    On Error Resume Next
    log.WriteLine tagx & " tcode=" & session.Info.Transaction & " prog=" & session.Info.Program & _
        " screen=" & session.Info.ScreenNumber & " title=[" & session.findById("wnd[0]").Text & "]" & _
        " sbar=[" & session.findById("wnd[0]/sbar").Text & "] sbarType=" & session.findById("wnd[0]/sbar").MessageType
    Err.Clear
    On Error GoTo 0
End Sub

Sub Finish(code)
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    If Exists("wnd[1]") Then
        log.WriteLine "popup on exit: [" & session.findById("wnd[1]").Text & "] - cancel"
        session.findById("wnd[1]").sendVKey 12
        WScript.Sleep 800
    End If
    Err.Clear
    log.WriteLine "exit code=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Echo "inspect_058_result " & TAG & " exit=" & code
    WScript.Quit code
End Sub

Sub DumpTree(comp, depth, lvl)
    Dim c, k, t, s
    On Error Resume Next
    t = comp.Type
    s = String(lvl * 2, " ") & t & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[") + 1) & " text=[" & comp.Text & "]"
    If t = "GuiShell" Then s = s & " sub=" & comp.SubType
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

' collect GridView shells anywhere under a component
Dim grids()
Dim ng
ng = 0
Sub FindGrids(comp)
    Dim c, k, t, st
    On Error Resume Next
    t = comp.Type
    st = ""
    If t = "GuiShell" Then st = comp.SubType
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

Sub DumpGrid(g, maxRows)
    Dim ord, c, r, line, vis, winTop, nr
    On Error Resume Next
    log.WriteLine "GRID " & Mid(g.Id, InStr(g.Id, "/wnd[") + 1) & " rows=" & g.RowCount & " cols=" & g.ColumnCount
    log.WriteLine "  toolbar buttons=" & g.ToolbarButtonCount
    For c = 0 To g.ToolbarButtonCount - 1
        log.WriteLine "    tb[" & c & "] id=" & g.GetToolbarButtonId(c) & " type=" & g.GetToolbarButtonType(c) & _
                      " enabled=" & g.GetToolbarButtonEnabled(c) & " text=[" & g.GetToolbarButtonText(c) & "]"
        Err.Clear
    Next
    Set ord = g.ColumnOrder
    line = ""
    For c = 0 To ord.Count - 1
        If c > 0 Then line = line & vbTab
        line = line & ord(c)
    Next
    log.WriteLine "  #COLS" & vbTab & line
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

' ---------- selection screen ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & TC
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2500
Err.Clear
On Error GoTo 0
Status "sel:"
If session.Info.ScreenNumber <> 1000 Then Finish 11

On Error Resume Next
If Exists("wnd[0]/usr/radPA_NEW") Then session.findById("wnd[0]/usr/radPA_NEW").Select
If Exists("wnd[0]/usr/radPA_DISP") Then session.findById("wnd[0]/usr/radPA_DISP").Select
If Exists("wnd[0]/usr/txtPA_SPMON") Then
    session.findById("wnd[0]/usr/txtPA_SPMON").Text = SPMON
Else
    session.findById("wnd[0]/usr/txtSO_SPMON-LOW").Text = SPMON
End If
If KUNNR <> "*" Then session.findById("wnd[0]/usr/ctxtSO_KUNNR-LOW").Text = KUNNR
If VKBUR <> "" Then session.findById("wnd[0]/usr/ctxtSO_VKBUR-LOW").Text = VKBUR
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
Status "sel after enter:"
If Exists("wnd[0]/usr/txtPA_SPMON") Then log.WriteLine "  PA_SPMON now=[" & session.findById("wnd[0]/usr/txtPA_SPMON").Text & "]"
If session.Info.ScreenNumber <> 1000 Then Finish 13

' ---------- execute (F8) ----------
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
Err.Clear
On Error GoTo 0
WScript.Sleep 2000
Status "after F8:"
If session.Info.ScreenNumber = 1000 And Not Exists("wnd[1]") Then
    On Error Resume Next
    If session.findById("wnd[0]/sbar").MessageType = "E" Then
        log.WriteLine "  F8 rejected. focused field=" & Mid(session.ActiveWindow.GuiFocus.Id, InStr(session.ActiveWindow.GuiFocus.Id, "/wnd[") + 1)
        Err.Clear
        On Error GoTo 0
        session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\SUBS_" & TAG & "_rejected.png", 2
        Finish 14
    End If
    Err.Clear
    On Error GoTo 0
End If
If Exists("wnd[1]") Then
    log.WriteLine "POPUP after F8: [" & session.findById("wnd[1]").Text & "]"
    DumpTree session.findById("wnd[1]"), 4, 0
    session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\SUBS_" & TAG & "_popup.png", 2
    session.findById("wnd[1]").sendVKey 12
    WScript.Sleep 1000
    Status "after popup cancel:"
End If

' ---------- result screen ----------
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\SUBS_" & TAG & ".png", 2
log.WriteLine "---- tbar[1] (GUI status)"
Dim b
For i = 0 To 50
    If Exists("wnd[0]/tbar[1]/btn[" & i & "]") Then
        Set b = session.findById("wnd[0]/tbar[1]/btn[" & i & "]")
        On Error Resume Next
        log.WriteLine "  btn[" & i & "] text=[" & b.Text & "] tip=[" & b.Tooltip & "] enabled=" & b.Changeable
        Err.Clear
        On Error GoTo 0
    End If
Next
log.WriteLine "---- wnd[0]/usr tree"
DumpTree session.findById("wnd[0]/usr"), 5, 0
If Exists("wnd[0]/shellcont") Then
    log.WriteLine "---- wnd[0]/shellcont tree"
    DumpTree session.findById("wnd[0]/shellcont"), 5, 0
End If
FindGrids session.findById("wnd[0]")
log.WriteLine "---- grids found=" & ng
For i = 0 To ng - 1
    DumpGrid grids(i), 60
Next

' ---------- optional 6th arg DRILL : select row 0 and press 상세리스트조회 (tbar[1] btn[43]) ----------
If WScript.Arguments.Count >= 6 Then
    If UCase(WScript.Arguments(5)) = "DRILL" And ng > 0 Then
        log.WriteLine "==== DRILL: select row 0, press tbar[1]/btn[43]"
        If Not Exists("wnd[0]/tbar[1]/btn[43]") Then
            log.WriteLine "  btn[43] not present"
            Finish 15
        End If
        On Error Resume Next
        grids(0).selectedRows = "0"
        WScript.Sleep 500
        session.findById("wnd[0]/tbar[1]/btn[43]").press
        WScript.Sleep 3000
        Err.Clear
        On Error GoTo 0
        Status "after drill:"
        If Exists("wnd[1]") Then
            log.WriteLine "POPUP after drill: [" & session.findById("wnd[1]").Text & "]"
            DumpTree session.findById("wnd[1]"), 4, 0
            session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\SUBS_" & TAG & "_drill_popup.png", 2
            session.findById("wnd[1]").sendVKey 12
            WScript.Sleep 1000
        End If
        session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\SUBS_" & TAG & "_drill.png", 2
        ng = 0
        FindGrids session.findById("wnd[0]")
        log.WriteLine "---- drill grids found=" & ng
        For i = 0 To ng - 1
            DumpGrid grids(i), 60
        Next
    End If
End If
' ---------- optional 6th arg FIPOS (case 4-4) : select row 0 and press tbar[1] btn[46] (batch FI posting) ----------
'   Expected to be BLOCKED. Any popup (confirmation or info) is logged and CANCELLED, never confirmed.
If WScript.Arguments.Count >= 6 Then
    If UCase(WScript.Arguments(5)) = "FIPOS" And ng > 0 Then
        log.WriteLine "==== FIPOS: select row 0, press tbar[1]/btn[46] tip=[" & session.findById("wnd[0]/tbar[1]/btn[46]").Tooltip & "]"
        On Error Resume Next
        grids(0).selectedRows = "0"
        WScript.Sleep 500
        session.findById("wnd[0]/tbar[1]/btn[46]").press
        WScript.Sleep 3000
        Err.Clear
        On Error GoTo 0
        Status "after FIPOS:"
        session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\SUBS_" & TAG & "_fipos.png", 2
        Dim fk
        For fk = 1 To 3
            If Not Exists("wnd[1]") Then Exit For
            log.WriteLine "POPUP after FIPOS: [" & session.findById("wnd[1]").Text & "] - CANCEL (not confirmed)"
            DumpTree session.findById("wnd[1]"), 4, 0
            session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\SUBS_" & TAG & "_fipos_popup" & fk & ".png", 2
            session.findById("wnd[1]").sendVKey 12
            WScript.Sleep 1500
            Status "after popup cancel:"
        Next
        ng = 0
        FindGrids session.findById("wnd[0]")
        log.WriteLine "---- grids after FIPOS=" & ng
        For i = 0 To ng - 1
            DumpGrid grids(i), 60
        Next
    End If
End If
Finish 0
