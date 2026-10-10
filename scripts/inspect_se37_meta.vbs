' inspect_se37_meta.vbs : READ-ONLY. SE37 test screen -> IT_DATA editor -> Metadata (btn[17]). Never executes the FM.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, c, k, lines(80), id, p, q, xy
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_se37_meta.txt", True, True)
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" And session.Info.SystemName <> "KDD" Then WScript.Quit 99

Sub Dump(title)
    For k = 0 To 80
        lines(k) = ""
    Next
    For Each c In session.findById("wnd[0]/usr").Children
        id = c.Id
        p = InStrRev(id, "[")
        q = InStrRev(id, "]")
        If p > 0 And q > p Then
            xy = Split(Mid(id, p + 1, q - p - 1), ",")
            If UBound(xy) = 1 Then
                If CInt(xy(1)) <= 80 And Trim(c.Text) <> "" Then lines(CInt(xy(1))) = lines(CInt(xy(1))) & " [" & xy(0) & "]" & Trim(c.Text)
            End If
        End If
    Next
    log.WriteLine "---- " & title & " prog=" & session.Info.Program & " screen=" & session.Info.ScreenNumber & " title=[" & session.findById("wnd[0]").Text & "] sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    For k = 0 To 80
        If lines(k) <> "" Then log.WriteLine "  y" & k & ":" & lines(k)
    Next
End Sub

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE37"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtRS38L-NAME").Text = "ZCRMF03_RFC_ZCRM_CHRG_CNFRM"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 3000
session.findById("wnd[0]/usr/lbl[34,10]").setFocus
session.findById("wnd[0]").sendVKey 2
WScript.Sleep 2000
Err.Clear
On Error GoTo 0
Dump "IT_DATA editor"
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[17]").press
WScript.Sleep 2500
Err.Clear
On Error GoTo 0
If connection.Children.Count > 0 Then
    On Error Resume Next
    Dim w
    Set w = session.findById("wnd[1]")
    If Err.Number = 0 Then
        Err.Clear
        log.WriteLine "POPUP [" & w.Text & "]"
        For Each c In w.findById("usr").Children
            log.WriteLine "   " & c.Type & " " & c.Id & " [" & c.Text & "]"
        Next
        w.HardCopy "C:\sap-test\results\shots\SE37_it_data_meta_popup.png", 2
        w.sendVKey 12
        WScript.Sleep 800
    End If
    Err.Clear
    On Error GoTo 0
End If
Dump "after Metadata"
session.findById("wnd[0]").HardCopy "C:\sap-test\results\shots\SE37_it_data_meta.png", 2
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1200
session.findById("wnd[1]").sendVKey 12
Err.Clear
On Error GoTo 0
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Quit 0