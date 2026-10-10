' inspect_deliv_popup.vbs : 세션/팝업 상태 확인 + '납품문서 생성' 대화상자 구조 덤프 + 안전한 취소 (실행 안 함)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i, s
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_deliv_popup.txt", True, True)

log.WriteLine "세션 수=" & connection.Children.Count
For i = 0 To connection.Children.Count - 1
    Set s = connection.Children.ElementAt(i)
    log.WriteLine "  " & s.Id & " tcode=" & s.Info.Transaction & " screen=" & s.Info.ScreenNumber
    On Error Resume Next
    log.WriteLine "    wnd[1]=" & s.findById("wnd[1]").Text
    Err.Clear
    On Error GoTo 0
Next

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

Sub Dump(comp, depth)
    Dim c, k, line
    If comp.Type = "GuiMenubar" Then Exit Sub
    line = String(depth * 2, " ") & comp.Type
    On Error Resume Next
    line = line & "/" & comp.SubType
    On Error GoTo 0
    line = line & " | " & comp.Id
    On Error Resume Next
    line = line & " | [" & Left(comp.Text, 70) & "]"
    line = line & " tip=[" & comp.Tooltip & "]"
    line = line & " chg=" & comp.Changeable
    On Error GoTo 0
    log.WriteLine line
    On Error Resume Next
    k = comp.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    On Error GoTo 0
    For c = 0 To k - 1
        Dump comp.Children.ElementAt(c), depth + 1
    Next
End Sub
Function HasPopup()
    On Error Resume Next
    HasPopup = False
    Dim w : Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then HasPopup = True
    Err.Clear
    On Error GoTo 0
End Function
Sub Status()
    log.WriteLine "  screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " sbar=[" & session.findById("wnd[0]/sbar").Text & "] popup=" & HasPopup()
End Sub
Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub

If Not HasPopup() Then
    log.WriteLine "=== 팝업 없음 → 1630 조회 후 CRE_DELIVERY 로 대화상자 열기 ==="
    Tcode "/nZSD600"
    With session
        .findById("wnd[0]/usr/radP_RBLIST").Select
        .findById("wnd[0]/usr/txtS_BSTKD-LOW").Text = "1630"
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "2026.01.01"
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = "2026.12.31"
        .findById("wnd[0]/usr/ctxtS_STAT1-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_STAT1-HIGH").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    WScript.Sleep 3500
    Dim gh : Set gh = session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell")
    gh.setCurrentCell 0, "BSTKD"
    gh.selectedRows = "0"
    gh.pressToolbarButton "CRE_DELIVERY"
    WScript.Sleep 2500
End If
Status
If HasPopup() Then
    log.WriteLine "=== 대화상자 전체 구조 ==="
    Dump session.findById("wnd[1]"), 0
    ' 안전한 취소 시도 순서: 1) 텍스트/툴팁에 취소·Cancel 있는 버튼  2) tbar[0]/btn[12]  3) F12  4) close
    Dim w1, c, comp, done, tb
    done = False
    Set w1 = session.findById("wnd[1]")
    On Error Resume Next
    Set tb = w1.findById("tbar[0]")
    If Err.Number = 0 Then
        For c = 0 To tb.Children.Count - 1
            Set comp = tb.Children.ElementAt(c)
            log.WriteLine "  tbar btn: " & comp.Id & " text=[" & comp.Text & "] tip=[" & comp.Tooltip & "]"
        Next
    End If
    Err.Clear
    On Error GoTo 0
    ' 1) 취소 버튼 찾기 (usr 내 GuiButton)
    Dim usr : Set usr = session.findById("wnd[1]/usr")
    For c = 0 To usr.Children.Count - 1
        Set comp = usr.Children.ElementAt(c)
        If comp.Type = "GuiButton" Then
            log.WriteLine "  usr btn: " & comp.Id & " text=[" & comp.Text & "] tip=[" & comp.Tooltip & "]"
            If Not done Then
                If InStr(comp.Text & comp.Tooltip, "취소") > 0 Or InStr(LCase(comp.Text & comp.Tooltip), "cancel") > 0 Then
                    log.WriteLine "  → 취소 버튼 누름: " & comp.Id
                    comp.press
                    WScript.Sleep 1000
                    done = Not HasPopup()
                End If
            End If
        End If
    Next
    If Not done Then
        On Error Resume Next
        Set comp = session.findById("wnd[1]/tbar[0]/btn[12]")
        If Err.Number = 0 Then
            log.WriteLine "  → tbar btn[12] 누름 text=[" & comp.Text & "] tip=[" & comp.Tooltip & "]"
            comp.press
            WScript.Sleep 1000
            done = Not HasPopup()
        End If
        Err.Clear
        On Error GoTo 0
    End If
    If Not done Then
        log.WriteLine "  → F12"
        session.findById("wnd[1]").sendVKey 12
        WScript.Sleep 1000
        done = Not HasPopup()
        If HasPopup() Then log.WriteLine "  F12 후 팝업: " & session.findById("wnd[1]").Text : Dump session.findById("wnd[1]"), 0
    End If
    If Not done Then
        log.WriteLine "  → wnd[1].close"
        session.findById("wnd[1]").close
        WScript.Sleep 1000
        done = Not HasPopup()
        If HasPopup() Then log.WriteLine "  close 후 팝업: " & session.findById("wnd[1]").Text : Dump session.findById("wnd[1]"), 0
    End If
    ' wnd[2] 확인 팝업(예/아니오)이 떴으면 '아니오'/취소
    On Error Resume Next
    Dim w2 : Set w2 = session.findById("wnd[2]")
    If Err.Number = 0 Then
        log.WriteLine "  wnd[2]: " & w2.Text
        Dump w2, 0
        w2.sendVKey 12
        WScript.Sleep 800
    End If
    Err.Clear
    On Error GoTo 0
    Status
    log.WriteLine "취소 결과: done=" & done
End If

If Not HasPopup() Then
    Tcode "/n"
    session.findById("wnd[0]").close
    log.WriteLine "세션 종료"
Else
    log.WriteLine "팝업이 남아 세션을 닫지 않음 ? 수동 확인 필요: " & session.Id
End If
log.Close
WScript.Echo "deliv popup done"
