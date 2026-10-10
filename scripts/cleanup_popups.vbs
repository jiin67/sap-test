' cleanup_popups.vbs : 열린 확인 팝업을 '아니오'로 안전 종료하고 여분 세션 정리 (절대 '예' 누르지 않음)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, i, guard
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\cleanup_popups.txt", True, True)

Function HasPopup(s)
    On Error Resume Next
    HasPopup = False
    Dim w : Set w = s.findById("wnd[1]")
    If Err.Number = 0 Then HasPopup = True
    Err.Clear
    On Error GoTo 0
End Function

' 팝업 하나를 안전하게 닫기: 아니오/No 버튼 → tbar btn[12] → close (예/Yes 는 절대 누르지 않음)
Sub SafeClose(s, wname)
    Dim w, usr, c, comp, pressed
    On Error Resume Next
    Set w = s.findById(wname)
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    log.WriteLine "  팝업 " & wname & " = [" & w.Text & "]"
    pressed = False
    Set usr = s.findById(wname & "/usr")
    If Err.Number = 0 Then
        For c = 0 To usr.Children.Count - 1
            Set comp = usr.Children.ElementAt(c)
            If comp.Type = "GuiButton" Then
                If InStr(comp.Text & comp.Tooltip, "아니오") > 0 Or InStr(LCase(comp.Text & comp.Tooltip), "no") > 0 Or InStr(comp.Text & comp.Tooltip, "취소") > 0 Then
                    log.WriteLine "    → 버튼 누름 " & comp.Id & " [" & comp.Text & "]"
                    comp.press : pressed = True : Exit For
                End If
            End If
        Next
    End If
    Err.Clear
    If Not pressed Then
        s.findById(wname & "/tbar[0]/btn[12]").press
        If Err.Number = 0 Then log.WriteLine "    → tbar btn[12]" : pressed = True
        Err.Clear
    End If
    If Not pressed Then
        s.findById(wname).close
        If Err.Number = 0 Then log.WriteLine "    → close" : pressed = True
        Err.Clear
    End If
    On Error GoTo 0
    WScript.Sleep 800
End Sub

log.WriteLine "세션 수=" & connection.Children.Count
For i = 0 To connection.Children.Count - 1
    Set session = connection.Children.ElementAt(i)
    log.WriteLine session.Id & " tcode=" & session.Info.Transaction & " screen=" & session.Info.ScreenNumber & " popup=" & HasPopup(session)
    guard = 0
    Do While HasPopup(session) And guard < 8
        SafeClose session, "wnd[3]"
        SafeClose session, "wnd[2]"
        SafeClose session, "wnd[1]"
        guard = guard + 1
    Loop
    log.WriteLine "  정리 후 popup=" & HasPopup(session) & " screen=" & session.Info.ScreenNumber
Next

' ses[0] 이 아닌 여분 세션은 /n 후 닫기 (ses[0] 은 사용자 세션이므로 건드리지 않음)
For i = connection.Children.Count - 1 To 1 Step -1
    Set session = connection.Children.ElementAt(i)
    If Not HasPopup(session) Then
        On Error Resume Next
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 600
        session.findById("wnd[0]").close
        WScript.Sleep 600
        log.WriteLine "여분 세션 닫음: index " & i
        Err.Clear
        On Error GoTo 0
    Else
        log.WriteLine "index " & i & " 팝업 남음 - 닫지 않음"
    End If
Next

log.WriteLine "최종 세션 수=" & connection.Children.Count
log.Close
WScript.Echo "cleanup done, sessions=" & connection.Children.Count
