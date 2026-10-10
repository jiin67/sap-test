' inspect_seldesc.vbs : READ-ONLY. Dump SE16N selection-field descriptions (SCRTEXT_M) per table.
'   args: <table> [<table> ...]   log: logs\inspect_seldesc.txt (UTF-16)
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, a
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
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_seldesc.txt", True, True)
n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500
If session.Info.SystemName <> "KDQ" Then WScript.Quit 99
Sub Go(tc)
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tc
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
End Sub

Sub ClosePopups()
    Dim w, k
    For k = 1 To 3
        On Error Resume Next
        Set w = session.findById("wnd[1]")
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        log.WriteLine "  POPUP wnd[1] title=[" & w.Text & "]"
        DumpTree w, 2, 0
        w.sendVKey 12
        WScript.Sleep 800
        Err.Clear
        On Error GoTo 0
    Next
End Sub

Sub DumpTree(comp, depth, lvl)
    Dim c, k, t, s
    On Error Resume Next
    t = comp.Type
    s = String(lvl * 2, " ") & t & " " & Mid(comp.Id, InStr(comp.Id, "/wnd[") + 1)
    s = s & " name=" & comp.Name & " text=[" & comp.Text & "]"
    If t = "GuiShell" Then s = s & " sub=" & comp.SubType
    If t = "GuiRadioButton" Or t = "GuiCheckBox" Then s = s & " sel=" & comp.Selected
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

Sub Status(tag)
    On Error Resume Next
    log.WriteLine tag & " tcode=" & session.Info.Transaction & " prog=" & session.Info.Program & _
        " screen=" & session.Info.ScreenNumber & " title=[" & session.findById("wnd[0]").Text & "]" & _
        " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    Err.Clear
    On Error GoTo 0
End Sub

' ---------- 3) SE16N selection descriptions ----------
Sub SelDesc(tbl)
    Dim r, pos, maxpos, txt, ed
    log.WriteLine "=================== SE16N " & tbl
    Go "SE16N"
    On Error Resume Next
    Set ed = session.findById("wnd[0]/usr/chkGD-EDIT")
    If Err.Number = 0 Then
        If ed.Selected Then
            log.WriteLine "  WARN edit checkbox was ON - clearing"
            ed.Selected = False
        End If
    End If
    Err.Clear
    session.findById("wnd[0]/usr/ctxtGD-TAB").Text = tbl
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 2500
    Err.Clear
    On Error GoTo 0
    Status "  after table:"
    ClosePopups
    On Error Resume Next
    maxpos = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VerticalScrollbar.Maximum
    If Err.Number <> 0 Then
        log.WriteLine "  no selection table: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    log.WriteLine "  scroll max=" & maxpos
    pos = 0
    Do
        On Error Resume Next
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VerticalScrollbar.Position = pos
        WScript.Sleep 400
        Err.Clear
        For r = 0 To 40
            txt = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & r & "]").Text
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            log.WriteLine "  pos " & pos & " row " & r & " : [" & txt & "]"
        Next
        On Error GoTo 0
        pos = pos + r
        If r = 0 Then Exit Do
    Loop While pos <= maxpos
End Sub


For a = 0 To WScript.Arguments.Count - 1
    SelDesc UCase(WScript.Arguments(a))
Next
Go ""
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_seldesc done"
WScript.Quit 0