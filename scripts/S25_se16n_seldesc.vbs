' S25_se16n_seldesc.vbs : READ-ONLY. List the SE16N selection-criterion descriptions (SCRTEXT_M) of a table,
'   so that S12 criteria ("<desc>=<value>") can be written with the right description.
'   args: <table>     log: logs\S25_<table>.txt (UTF-16)     Pure ASCII source.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, i, n0, TBL
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = Nothing
For i = 0 To application.Children.Count - 1
    Set base = application.Children.ElementAt(i).Children.ElementAt(0)
    If base.Info.SystemName = "KDQ" Then
        Set connection = application.Children.ElementAt(i)
        Exit For
    End If
Next
If connection Is Nothing Then WScript.Quit 99
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" Then WScript.Quit 99
If WScript.Arguments.Count < 1 Then WScript.Quit 1
TBL = UCase(WScript.Arguments(0))

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\S25_" & TBL & ".txt", True, True)
log.WriteLine "table=" & TBL

n0 = connection.Children.Count
base.createSession
For i = 1 To 40
    WScript.Sleep 500
    If connection.Children.Count > n0 Then Exit For
Next
Set session = connection.Children.ElementAt(connection.Children.Count - 1)
WScript.Sleep 1500

On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2000
session.findById("wnd[0]/usr/ctxtGD-TAB").Text = TBL
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2500
Dim ed
Set ed = session.findById("wnd[0]/usr/chkGD-EDIT")
If Err.Number = 0 Then
    If ed.Selected Then
        log.WriteLine "WARN edit checkbox was ON - clearing"
        ed.Selected = False
    End If
End If
Err.Clear

Dim tc, maxpos, pos, r, txt, fld, seen
Set tc = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC")
maxpos = tc.VerticalScrollbar.Maximum
If Err.Number <> 0 Then
    log.WriteLine "ERR selection table " & Err.Description
    Err.Clear
    maxpos = 0
End If
Set seen = CreateObject("Scripting.Dictionary")
pos = 0
Do
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VerticalScrollbar.Position = pos
    WScript.Sleep 400
    Err.Clear
    For r = 0 To 30
        txt = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-SCRTEXT_M[0," & r & "]").Text
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        If Not seen.Exists(txt) Then
            seen.Add txt, 1
            log.WriteLine "  [" & txt & "]"
        End If
    Next
    pos = pos + 15
Loop While pos <= maxpos

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
Err.Clear
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
