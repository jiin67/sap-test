' S23_se11_field.vbs : READ-ONLY. Show the DDIC field list of one table in SE11 and find given fields.
'   args: <tag> <TABLE> <FIELD>[,<FIELD>...]
'   Scrolls the SE11 field table control and logs every row; marks the wanted fields and shots the page they are on.
'   log: logs\S23_<tag>.txt (UTF-16)   shots: results\shots\S23_<tag>_<field>.png    exit 0 ok / 20 not found / 99 guard
'   Pure ASCII source (Korean only via ChrW).
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Dim TAG, TBL, WANT, wanted, found, pos, rowsSeen, tcId, r, nm, dt, ln, ky, hit, total

If WScript.Arguments.Count < 3 Then
    WScript.Echo "usage: S23_se11_field.vbs <tag> <TABLE> <FIELD>[,<FIELD>]"
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
TBL = UCase(WScript.Arguments(1))
WANT = UCase(WScript.Arguments(2))
wanted = Split(WANT, ",")

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
Set log = fso.CreateTextFile("C:\sap-test\logs\S23_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG & " table=" & TBL & " want=" & WANT
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
    WScript.Echo "S23 " & TAG & " exit=" & code
    WScript.Quit code
End Sub

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE11"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/radRSRD1-TBMA").Select
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").Text = TBL
session.findById("wnd[0]").sendVKey 7
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
log.WriteLine "screen=" & session.Info.ScreenNumber & " title=[" & T("wnd[0]") & "]"

tcId = "wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0"
If Not Exists(tcId) Then
    log.WriteLine "FAIL no field table control"
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
    WScript.Sleep 500
    Err.Clear
    On Error GoTo 0
    rowsSeen = 0
    hit = False
    For r = 0 To 30
        nm = ""
        On Error Resume Next
        nm = session.findById(tcId & "/txtDD03D-FIELDNAME[0," & r & "]").Text
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        ky = session.findById(tcId & "/chkDD03P-KEYFLAG[1," & r & "]").Selected
        dt = session.findById(tcId & "/txtDD03D-DATATYPE[4," & r & "]").Text
        ln = session.findById(tcId & "/txtDD03P-LENG[5," & r & "]").Text & " ro=" & session.findById(tcId & "/ctxtDD03D-ROLLNAME[3," & r & "]").Text & " txt=[" & session.findById(tcId & "/txtDD03P-DDTEXT[7," & r & "]").Text & "]"
        If Err.Number <> 0 Then
            Err.Clear
            dt = ""
        End If
        On Error GoTo 0
        rowsSeen = rowsSeen + 1
        If Trim(nm) <> "" Then
            log.WriteLine "  " & (pos + r) & " " & nm & " key=" & ky & " type=" & dt & " len=" & ln
            For i = 0 To UBound(wanted)
                If UCase(Trim(nm)) = Trim(wanted(i)) Then
                    log.WriteLine "  *** FOUND " & nm & " type=" & dt & " len=" & ln & " key=" & ky
                    found = found + 1
                    hit = True
                End If
            Next
        End If
    Next
    If hit Then
        On Error Resume Next
        session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S23_" & TAG & "_" & pos & ".png", 2
        Err.Clear
        On Error GoTo 0
    End If
    If rowsSeen = 0 Then Exit Do
    pos = pos + rowsSeen
Loop While pos <= total

log.WriteLine "found=" & found & " of " & (UBound(wanted) + 1)
If found < UBound(wanted) + 1 Then Finish 20
Finish 0
