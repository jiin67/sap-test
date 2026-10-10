' S24_se38_textsym.vbs : READ-ONLY. List the text symbols of a report (SE38 -> 텍스트요소 -> 텍스트기호) and look for one.
'   args: <tag> <PROGRAM> <SYMBOL>
'   log: logs\S24_<tag>.txt (UTF-16)   exit 0 found / 20 not found / 99 guard.  Pure ASCII source.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Dim TAG, PRG, SYM, tcId, pos, r, nm, tx, rowsSeen, total, found

If WScript.Arguments.Count < 3 Then
    WScript.Echo "usage: S24_se38_textsym.vbs <tag> <PROGRAM> <SYMBOL>"
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
PRG = UCase(WScript.Arguments(1))
SYM = UCase(WScript.Arguments(2))

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
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
Set log = fso.CreateTextFile("C:\sap-test\logs\S24_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG & " program=" & PRG & " symbol=" & SYM
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
If session.Info.SystemName <> "KDQ" Then WScript.Quit 99
If session.Info.Transaction <> "SESSION_MANAGER" And session.Info.Transaction <> "SMEN" Then
    log.WriteLine "ABORT: picked session is busy (" & session.Info.Transaction & ")"
    log.Close
    WScript.Quit 14
End If

Function Exists(id)
    Dim o
    On Error Resume Next
    Set o = session.findById(id)
    Exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Sub Finish(code)
    On Error Resume Next
    If Exists("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 800
    Err.Clear
    On Error GoTo 0
    log.WriteLine "exit code=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Echo "S24 " & TAG & " exit=" & code
    WScript.Quit code
End Sub

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = PRG
session.findById("wnd[0]/usr/radRS38M-FUNC_TEXT").Select
session.findById("wnd[0]").sendVKey 7
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
log.WriteLine "screen=" & session.Info.ScreenNumber

tcId = "wnd[0]/usr/tabsTX_TABSTR_CONTROL/tabpIIII/ssubSCREEN_HEADER:SAPLSETXP:1210/tblSAPLSETXPTSYMBOL"
If Not Exists(tcId) Then
    log.WriteLine "FAIL no text symbol table control"
    Finish 20
End If
On Error Resume Next
total = session.findById(tcId).VerticalScrollbar.Maximum
Err.Clear
On Error GoTo 0
log.WriteLine "scrollbar max=" & total

found = 0
pos = 0
Do
    On Error Resume Next
    session.findById(tcId).VerticalScrollbar.Position = pos
    WScript.Sleep 400
    Err.Clear
    On Error GoTo 0
    rowsSeen = 0
    For r = 0 To 40
        nm = ""
        On Error Resume Next
        nm = session.findById(tcId & "/txtRS38M-INUM[0," & r & "]").Text
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        tx = session.findById(tcId & "/txtRS38M-ITEXT[1," & r & "]").Text
        If Err.Number <> 0 Then
            Err.Clear
            tx = ""
        End If
        On Error GoTo 0
        rowsSeen = rowsSeen + 1
        If Trim(nm) <> "" Then
            log.WriteLine "  " & (pos + r) & " [" & nm & "] " & tx
            If UCase(Trim(nm)) = SYM Then
                log.WriteLine "  *** FOUND " & nm & " = [" & tx & "]"
                found = found + 1
                On Error Resume Next
                session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S24_" & TAG & "_" & SYM & ".png", 2
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next
    If rowsSeen = 0 Then Exit Do
    pos = pos + rowsSeen
Loop While pos <= total

log.WriteLine "found=" & found
If found = 0 Then Finish 20
Finish 0
