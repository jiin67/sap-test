'==========================================================
'  SAP GUI Scripting 활성화 진단 스크립트
'  사용법 : SAP 에 로그인한 상태에서 이 파일을 더블클릭
'==========================================================

Option Explicit

Dim SapGuiAuto, app, connection, session
Dim msg, okcd

msg = "=== SAP GUI Scripting 진단 결과 ===" & vbCrLf & vbCrLf

On Error Resume Next

'---------- 1. SAP GUI 접근 (클라이언트 설정 확인) ----------
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    msg = msg & "[1] SAP GUI 접근 ......... 실패" & vbCrLf & vbCrLf
    msg = msg & "> 원인" & vbCrLf
    msg = msg & "   SAP GUI 가 실행 중이 아니거나," & vbCrLf
    msg = msg & "   클라이언트 스크립팅 설정이 아직 적용되지 않았습니다." & vbCrLf & vbCrLf
    msg = msg & "> 조치" & vbCrLf
    msg = msg & "   1) 모든 SAP 세션을 닫습니다" & vbCrLf
    msg = msg & "   2) 작업표시줄 우측 트레이 아이콘에서도 SAP Logon 종료" & vbCrLf
    msg = msg & "   3) SAP Logon 재실행 후 로그인, 다시 이 파일 실행"
    MsgBox msg, vbCritical, "SAP Scripting 진단"
    WScript.Quit
End If
msg = msg & "[1] SAP GUI 접근 ......... 정상" & vbCrLf
Err.Clear

'---------- 2. 스크립팅 엔진 (서버 파라미터 확인) ----------
Set app = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Then
    msg = msg & "[2] 스크립팅 엔진 ........ 실패" & vbCrLf & vbCrLf
    msg = msg & "> 원인" & vbCrLf
    msg = msg & "   서버에서 스크립팅이 차단된 상태입니다." & vbCrLf
    msg = msg & "   sapgui/user_scripting = FALSE" & vbCrLf & vbCrLf
    msg = msg & "> 조치" & vbCrLf
    msg = msg & "   Basis 팀에 파라미터 변경을 요청하세요." & vbCrLf
    msg = msg & "   (RZ11 즉시 변경 + RZ10 프로파일 영구 반영)"
    MsgBox msg, vbCritical, "SAP Scripting 진단"
    WScript.Quit
End If
msg = msg & "[2] 스크립팅 엔진 ........ 정상" & vbCrLf
Err.Clear

'---------- 3. 활성 세션 확인 ----------
If app.Children.Count = 0 Then
    msg = msg & "[3] 활성 세션 ............ 없음" & vbCrLf & vbCrLf
    msg = msg & "> 조치" & vbCrLf
    msg = msg & "   SAP 에 로그인한 뒤 다시 실행하세요."
    MsgBox msg, vbExclamation, "SAP Scripting 진단"
    WScript.Quit
End If

Set connection = app.Children(0)
Set session = connection.Children(0)
If Err.Number <> 0 Then
    msg = msg & "[3] 활성 세션 ............ 실패" & vbCrLf & vbCrLf
    msg = msg & "> 조치" & vbCrLf
    msg = msg & "   SAP 에 로그인한 뒤 다시 실행하세요."
    MsgBox msg, vbExclamation, "SAP Scripting 진단"
    WScript.Quit
End If

msg = msg & "[3] 활성 세션 ............ 정상" & vbCrLf
msg = msg & "      시스템   : " & session.Info.SystemName & vbCrLf
msg = msg & "      클라이언트 : " & session.Info.Client & vbCrLf
msg = msg & "      사용자   : " & session.Info.User & vbCrLf
Err.Clear

'---------- 4. 쓰기 권한 (readonly 파라미터 확인) ----------
' 명령 입력창에 빈 문자열을 넣어보는 무해한 테스트입니다.
Set okcd = session.findById("wnd[0]/tbar[0]/okcd")
okcd.Text = ""
If Err.Number <> 0 Then
    msg = msg & "[4] 쓰기 권한 ............ 차단" & vbCrLf & vbCrLf
    msg = msg & "> 원인" & vbCrLf
    msg = msg & "   sapgui/user_scripting_set_readonly = TRUE" & vbCrLf
    msg = msg & "   화면 조회는 되지만 입력·실행은 불가능합니다." & vbCrLf & vbCrLf
    msg = msg & "> 조치" & vbCrLf
    msg = msg & "   Basis 팀에 FALSE 로 변경을 요청하세요."
    MsgBox msg, vbExclamation, "SAP Scripting 진단"
    WScript.Quit
End If
msg = msg & "[4] 쓰기 권한 ............ 정상" & vbCrLf & vbCrLf
Err.Clear

msg = msg & "-------------------------------" & vbCrLf
msg = msg & "모든 항목 통과. 자동화를 시작할 수 있습니다." & vbCrLf
msg = msg & "Alt+F12 > Script Recording and Playback 으로" & vbCrLf
msg = msg & "원하는 동작을 녹화해 코드를 뽑아내세요."

MsgBox msg, vbInformation, "SAP Scripting 진단"
