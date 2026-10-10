' inspect_chgmat_tbar.vbs : 자재변경 대화상자의 tbar 버튼 + 확인 트리거 동작 파악 (저장 안 함)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_chgmat_tbar.txt", True, True)
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
Const HG = "wnd[0]/shellcont/shell/shellcont[0]/shell"
Const IG = "wnd[0]/shellcont/shell/shellcont[1]/shell"

Sub Tcode(t)
    session.findById("wnd[0]/tbar[0]/okcd").Text = t
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1200
End Sub
Function Sbar()
    Sbar = session.findById("wnd[0]/sbar").Text
End Function
Function HasPopup()
    On Error Resume Next
    HasPopup = False
    Dim w : Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then HasPopup = True
    Err.Clear
    On Error GoTo 0
End Function
' 확인 팝업(예/아니오) 아니오로 닫기
Sub CancelConfirm(wname)
    Dim usr, c, comp
    On Error Resume Next
    Set usr = session.findById(wname & "/usr")
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    For c = 0 To usr.Children.Count - 1
        Set comp = usr.Children.ElementAt(c)
        If comp.Type = "GuiButton" Then
            If InStr(comp.Text & comp.Tooltip, "아니오") > 0 Then comp.press : WScript.Sleep 700 : Exit Sub
        End If
    Next
    On Error GoTo 0
End Sub

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
session.findById(HG).setCurrentCell 0, "BSTKD"
session.findById(HG).selectedRows = "0"
On Error Resume Next
session.findById(HG).doubleClickCurrentCell
Err.Clear
On Error GoTo 0
WScript.Sleep 1500
session.findById(IG).setCurrentCell 0, "MATNR"
session.findById(IG).selectedRows = "0"
session.findById(IG).pressToolbarButton "CHANGE_MAT"
WScript.Sleep 1500
log.WriteLine "팝업=" & session.findById("wnd[1]").Text
' tbar[0] 버튼 나열
Dim tb, c, comp
Set tb = session.findById("wnd[1]/tbar[0]")
log.WriteLine "tbar[0] children=" & tb.Children.Count
For c = 0 To tb.Children.Count - 1
    Set comp = tb.Children.ElementAt(c)
    log.WriteLine "  " & comp.Id & " type=" & comp.Type & " text=[" & comp.Text & "] tip=[" & comp.Tooltip & "]"
Next

' 빈 자재 상태로 확인 버튼(tbar btn[0]) 눌러 검증 메시지 유도
log.WriteLine "--- 빈 자재 + tbar[0]/btn[0] ---"
On Error Resume Next
session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = ""
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 1500
log.WriteLine "  sbar=[" & Sbar() & "] popup2=" & (Not (session.findById("wnd[2]").Text = "" And Err.Number <> 0))
log.WriteLine "  wnd[1]=[" & session.findById("wnd[1]").Text & "]"
Err.Clear
' wnd[1] 상태바
log.WriteLine "  wnd[1]/sbar=[" & session.findById("wnd[1]/sbar").Text & "]"
Err.Clear
On Error GoTo 0

' 미존재 자재 + 확인
log.WriteLine "--- 미존재 자재 ZZTEST-NOTEXIST + btn[0] ---"
On Error Resume Next
session.findById("wnd[1]/usr/txtGS_SCREEN_0110-MATNR").Text = "ZZTEST-NOTEXIST"
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 1500
log.WriteLine "  sbar=[" & Sbar() & "] wnd1sbar=[" & session.findById("wnd[1]/sbar").Text & "] wnd1=[" & session.findById("wnd[1]").Text & "]"
Err.Clear
On Error GoTo 0

' 열려있는 팝업/확인 정리 (아니오/close)
Dim g
For g = 1 To 6
    If Not HasPopup() Then Exit For
    CancelConfirm "wnd[2]"
    CancelConfirm "wnd[1]"
    On Error Resume Next
    session.findById("wnd[1]").close
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 700
Next
log.WriteLine "최종 popup=" & HasPopup() & " screen=" & session.Info.ScreenNumber

If Not HasPopup() Then
    Tcode "/n"
    session.findById("wnd[0]").close
End If
log.Close
WScript.Echo "chgmat_tbar done popup=" & HasPopup()
