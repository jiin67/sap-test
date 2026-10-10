' S14_se37_rfc.vbs : call ZCRMF03_RFC_ZCRM_CHRG_CNFRM through the SE37 single test and record the result.
'   User-approved (2026-09-11) for scenario groups 2/3/5 with the designated CSLTNO only.
'   args: <tag> <IV_ACTION: S | C | X | _ (= blank)> [<CSLTNO>:<YYYYMMDD or empty>] ...
'         no item args = IT_DATA left empty (case 2-8)
'   Before F8 the script checks that IV_ACTION and the IT_DATA entry count are what was asked;
'   otherwise it leaves without executing (exit 12).
'   log: logs\S14_<tag>.txt (UTF-16), shots: results\shots\S14_<tag>_*.png
'   exit: 0 executed / 11 not on test screen / 12 input check failed / 13 no result screen
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0
Dim TAG, ACT, nItems, items()

If WScript.Arguments.Count < 2 Then
    WScript.Echo "usage: S14_se37_rfc.vbs <tag> <S|C|X|_> [csltno:yyyymmdd] ..."
    WScript.Quit 1
End If
TAG = WScript.Arguments(0)
ACT = WScript.Arguments(1)
If ACT = "_" Then ACT = ""
nItems = WScript.Arguments.Count - 2
If nItems > 0 Then
    ReDim items(nItems - 1)
    For i = 0 To nItems - 1
        items(i) = Split(WScript.Arguments(i + 2) & ":", ":")
    Next
End If

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
Set log = fso.CreateTextFile("C:\sap-test\logs\S14_" & TAG & ".txt", True, True)
log.WriteLine "tag=" & TAG & " IV_ACTION=[" & ACT & "] items=" & nItems
For i = 0 To nItems - 1
    log.WriteLine "  IT_DATA[" & (i + 1) & "] CSLTNO=[" & items(i)(0) & "] CFDATE1=[" & items(i)(1) & "]"
Next

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" Then WScript.Quit 99
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

Sub Shot(name)
    On Error Resume Next
    session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\S14_" & TAG & "_" & name & ".png", 2
    Err.Clear
    On Error GoTo 0
End Sub

Sub Finish(code)
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
    If Exists("wnd[1]") Then
        log.WriteLine "popup on exit [" & T("wnd[1]") & "] - cancel"
        session.findById("wnd[1]").sendVKey 12
        WScript.Sleep 800
    End If
    Err.Clear
    On Error GoTo 0
    log.WriteLine "exit code=" & code
    log.Close
    session.findById("wnd[0]").close
    WScript.Echo "S14 " & TAG & " exit=" & code
    WScript.Quit code
End Sub

' dump a list screen line by line: "y: [x]text | [x]text"
Sub DumpList(title)
    Dim c, id, p, q, xy, x, y, lines(), maxy, k
    maxy = 60
    ReDim lines(maxy)
    For Each c In session.findById("wnd[0]/usr").Children
        id = c.Id
        p = InStrRev(id, "[")
        q = InStrRev(id, "]")
        If p > 0 And q > p Then
            xy = Split(Mid(id, p + 1, q - p - 1), ",")
            If UBound(xy) = 1 Then
                x = CInt(xy(0))
                y = CInt(xy(1))
                If y <= maxy Then
                    If Trim(c.Text) <> "" Then lines(y) = lines(y) & " [" & x & "]" & Trim(c.Text)
                End If
            End If
        End If
    Next
    log.WriteLine "---- " & title
    For k = 0 To maxy
        If lines(k) <> "" Then log.WriteLine "  y" & k & ":" & lines(k)
    Next
End Sub

' find the label on the list whose text is exactly <name>; return its y (or -1)
Function FindLineY(name)
    Dim c, id, p, q, xy
    FindLineY = -1
    For Each c In session.findById("wnd[0]/usr").Children
        If Trim(c.Text) = name Then
            id = c.Id
            p = InStrRev(id, "[")
            q = InStrRev(id, "]")
            xy = Split(Mid(id, p + 1, q - p - 1), ",")
            FindLineY = CInt(xy(1))
            Exit Function
        End If
    Next
End Function

' leading integer of a label like "1 Entry" / "3 엔트리" ("" if none)
Function LeadNum(s)
    Dim m, re
    Set re = New RegExp
    re.Pattern = "^\s*(\d+)"
    LeadNum = ""
    If re.Test(s) Then
        Set m = re.Execute(s)
        LeadNum = m(0).SubMatches(0)
    End If
End Function

' ---------- open the single-test screen ----------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE37"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtRS38L-NAME").Text = "ZCRMF03_RFC_ZCRM_CHRG_CNFRM"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 3000
Err.Clear
On Error GoTo 0
Status "test screen:"
If session.Info.Program <> "SAPMSSY0" Or Not Exists("wnd[0]/usr/txt[34,9]") Or T("wnd[0]/usr/lbl[2,9]") <> "IV_ACTION" Then Finish 11

' ---------- IV_ACTION ----------
session.findById("wnd[0]/usr/txt[34,9]").Text = ACT

' ---------- IT_DATA ----------
If nItems > 0 Then
    session.findById("wnd[0]/usr/lbl[34,10]").setFocus
    session.findById("wnd[0]").sendVKey 2
    WScript.Sleep 2000
    Status "editor:"
    If InStr(T("wnd[0]"), "IT_DATA") = 0 Then
        log.WriteLine "  not in IT_DATA editor"
        Finish 12
    End If
    Dim y, k
    For i = 0 To nItems - 1
        y = 6 + i
        For k = 1 To 2
            If Exists("wnd[0]/usr/txt[1," & y & "]") Then Exit For
            session.findById("wnd[0]/tbar[1]/btn[32]").press
            WScript.Sleep 1000
        Next
        If Not Exists("wnd[0]/usr/txt[1," & y & "]") Then
            log.WriteLine "  no input line for item " & (i + 1)
            DumpList "editor (no line)"
            Finish 12
        End If
        session.findById("wnd[0]/usr/txt[1," & y & "]").Text = items(i)(0)
        session.findById("wnd[0]/usr/txt[15," & y & "]").Text = items(i)(1)
    Next
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1000
    DumpList "editor after input"
    Shot "1_it_data"
    session.findById("wnd[0]").sendVKey 3
    WScript.Sleep 1500
    Status "back on test screen:"
End If

' ---------- verify the input before executing ----------
Dim cnt, actNow
actNow = Trim(T("wnd[0]/usr/txt[34,9]"))
cnt = LeadNum(T("wnd[0]/usr/lbl[37,10]"))   ' "0 엔트리" / "1 Entry" / "n 엔트리"
log.WriteLine "check: IV_ACTION now=[" & actNow & "] IT_DATA entries=[" & cnt & "]"
DumpList "test screen before F8"
If actNow <> ACT Or cnt <> CStr(nItems) Then
    log.WriteLine "  input check FAILED - not executing"
    Finish 12
End If
Shot "2_before_exec"

' ---------- execute ----------
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 4000
Status "after F8:"
If Exists("wnd[1]") Then
    log.WriteLine "POPUP after F8 [" & T("wnd[1]") & "]"
    session.findById("wnd[1]").HardCopy "C:\sap-test\results\shots\S14_" & TAG & "_popup.png", 2
    session.findById("wnd[1]").sendVKey 12
    WScript.Sleep 1000
End If
DumpList "result screen"
Shot "3_result"

' ---------- open ET_RETURN ----------
y = FindLineY("ET_RETURN")
log.WriteLine "ET_RETURN line y=" & y
If y < 0 Then Finish 13
On Error Resume Next
If Exists("wnd[0]/usr/lbl[34," & y & "]") Then
    session.findById("wnd[0]/usr/lbl[34," & y & "]").setFocus
Else
    session.findById("wnd[0]/usr/lbl[2," & y & "]").setFocus
End If
session.findById("wnd[0]").sendVKey 2
WScript.Sleep 2000
Err.Clear
On Error GoTo 0
Status "ET_RETURN view:"
DumpList "ET_RETURN"
Shot "4_et_return"
Finish 0
