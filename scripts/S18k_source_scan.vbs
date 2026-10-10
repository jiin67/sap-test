' KDD variant of S18_source_scan.vbs (2026-10-03): same script, system guard KDD. READ-ONLY.
' S18_source_scan.vbs : READ-ONLY. Run RS_ABAP_SOURCE_SCAN (via SE38) to read ABAP source lines.
'   args: <tag> FIELDS
'           -> dump the selection screen fields of RS_ABAP_SOURCE_SCAN (to learn field ids)
'         <tag> SCAN <search string> <fieldid>=<value> [<fieldid>=<value> ...]
'           -> fill txtSSTRING-LOW with the search string, set the given fields (e.g. program name),
'              tick chkINCLU (explode includes), execute, dump the hit list line by line.
'   Nothing is changed in the system. log: logs\S18k_<tag>.txt (UTF-16)
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, TAG, MODE

If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: S18_source_scan.vbs <tag> FIELDS | <tag> SCAN <string> <fieldid>=<value> ..."
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
MODE = UCase(WScript.Arguments(1))

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
' pick the KDQ connection explicitly (KDP/KDD may be open in the same SAP GUI)
Set connection = Nothing
For i = 0 To application.Children.Count - 1
    If application.Children.ElementAt(i).Children.ElementAt(0).Info.SystemName = "KDD" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then WScript.Quit 99
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDD" Then WScript.Quit 99

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S18k_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG & " mode=" & MODE

If connection.Children.Count >= 6 Then
    log.WriteLine "ABORT: session limit"
    log.Close
    WScript.Quit 14
End If
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDD" Then WScript.Quit 99
If session.Info.Transaction <> "SESSION_MANAGER" And session.Info.Transaction <> "SMEN" Then
    log.WriteLine "ABORT: picked session is busy (" & session.Info.Transaction & ")"
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
    log.WriteLine x & " prog=" & session.Info.Program & " screen=" & session.Info.ScreenNumber & _
        " title=[" & T("wnd[0]") & "] sbar=[" & T("wnd[0]/sbar") & "]"
End Sub

Sub DumpTree(comp, depth, lvl)
    Dim c, k, tp, s
    On Error Resume Next
    tp = comp.Type
    s = String(lvl * 2, " ") & tp & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[") + 1) & " text=[" & comp.Text & "]"
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

' list screen -> "y: [x]text [x]text"
Sub DumpList(maxy)
    Dim c, id, p, q, xy, lines(), k
    ReDim lines(maxy)
    For Each c In session.findById("wnd[0]/usr").Children
        id = c.Id
        p = InStrRev(id, "[")
        q = InStrRev(id, "]")
        If p > 0 And q > p Then
            xy = Split(Mid(id, p + 1, q - p - 1), ",")
            If UBound(xy) = 1 Then
                If CInt(xy(1)) <= maxy And Trim(c.Text) <> "" Then lines(CInt(xy(1))) = lines(CInt(xy(1))) & " [" & xy(0) & "]" & c.Text
            End If
        End If
    Next
    For k = 0 To maxy
        If lines(k) <> "" Then log.WriteLine "  y" & k & ":" & lines(k)
    Next
End Sub

Sub DumpListAbs(maxy, off)
    Dim c, id, p, q, xy, lines(), k
    ReDim lines(maxy)
    For Each c In session.findById("wnd[0]/usr").Children
        id = c.Id
        p = InStrRev(id, "[")
        q = InStrRev(id, "]")
        If p > 0 And q > p Then
            xy = Split(Mid(id, p + 1, q - p - 1), ",")
            If UBound(xy) = 1 Then
                If CInt(xy(1)) <= maxy And Trim(c.Text) <> "" Then lines(CInt(xy(1))) = lines(CInt(xy(1))) & " [" & xy(0) & "]" & c.Text
            End If
        End If
    Next
    For k = 3 To maxy
        If lines(k) <> "" Then log.WriteLine "  a" & (off + k) & ":" & lines(k)
    Next
End Sub

' text of the last non-empty label on the current list page (to detect the end of paging)
Function LastLineText()
    Dim c, s
    s = ""
    On Error Resume Next
    For Each c In session.findById("wnd[0]/usr").Children
        If Trim(c.Text) <> "" Then s = c.Id & "|" & c.Text
    Next
    Err.Clear
    On Error GoTo 0
    LastLineText = s
End Function

Sub Finish(code)
    On Error Resume Next
    If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1000
    If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
    Err.Clear
    On Error GoTo 0
    log.WriteLine "exit code=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Echo "S18 " & TAG & " exit=" & code
    WScript.Quit code
End Sub

' ---------- start RS_ABAP_SOURCE_SCAN through SE38 ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "RS_ABAP_SOURCE_SCAN"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
Status "selection:"

If MODE = "FIELDS" Then
    DumpTree session.findById("wnd[0]/usr"), 3, 0
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S18k_" & TAG & "_selection.png", 2
    Finish 0
End If

Dim a, eqp
On Error Resume Next
session.findById("wnd[0]/usr/txtSSTRING-LOW").Text = WScript.Arguments(2)
For a = 3 To WScript.Arguments.Count - 1
    eqp = InStr(WScript.Arguments(a), "=")
    If Left(WScript.Arguments(a), 3) = "chk" Then
        session.findById("wnd[0]/usr/" & Left(WScript.Arguments(a), eqp - 1)).Selected = (Mid(WScript.Arguments(a), eqp + 1) = "X")
    Else
        session.findById("wnd[0]/usr/" & Left(WScript.Arguments(a), eqp - 1)).Text = Mid(WScript.Arguments(a), eqp + 1)
    End If
    log.WriteLine "  set " & WScript.Arguments(a)
    If Err.Number <> 0 Then
        log.WriteLine "  ERR " & Err.Description
        Err.Clear
    End If
Next
If Exists("wnd[0]/usr/chkINCLU") Then session.findById("wnd[0]/usr/chkINCLU").Selected = True
Err.Clear
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 6000
Err.Clear
On Error GoTo 0
Status "result:"
If Exists("wnd[1]") Then
    log.WriteLine "POPUP [" & T("wnd[1]") & "]"
    DumpTree session.findById("wnd[1]"), 4, 0
End If
' page through the hit list by scrollbar position (only visible rows are exposed; overlap, absolute row index)
Dim pos, prevPos, vsb
pos = 0
prevPos = -1
Do
    On Error Resume Next
    Set vsb = session.findById("wnd[0]/usr").VerticalScrollbar
    vsb.Position = pos
    WScript.Sleep 300
    pos = vsb.Position
    Err.Clear
    On Error GoTo 0
    If pos = prevPos Then Exit Do
    log.WriteLine "  ---- pos " & pos
    DumpListAbs 300, pos
    If prevPos = -1 Then
        On Error Resume Next
        session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S18k_" & TAG & "_result.png", 2
        If Err.Number <> 0 Then log.WriteLine "  HardCopy err " & Err.Number
        Err.Clear
        On Error GoTo 0
    End If
    prevPos = pos
    pos = pos + 20
Loop
Finish 0
