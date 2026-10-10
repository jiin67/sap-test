' inspect_scan.vbs : RPR_ABAP_SOURCE_SCAN ???? Z33SDR600(+INCLUDE) ??? ??? (?ฌา?????)
Option Explicit
Dim SapGuiAuto, application, connection, session, base, fso, log, n0, i
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then WScript.Quit 99
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_scan.txt", True, True)
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

' ????? ????? ???? ?????? ?????? ????
Sub DumpList()
    Dim page, usr, c, comp, line, prev, cur, rows, nm, p1, p2, col, row, maxrow
    prev = ""
    For page = 1 To 60
        Set usr = session.findById("wnd[0]/usr")
        cur = ""
        Dim d : Set d = CreateObject("Scripting.Dictionary")
        maxrow = 0
        For c = 0 To usr.Children.Count - 1
            Set comp = usr.Children.ElementAt(c)
            If comp.Type = "GuiLabel" Then
                nm = comp.Id
                nm = Mid(nm, InStrRev(nm, "/") + 1)
                p1 = InStr(nm, "[") : p2 = InStr(nm, ",")
                If p1 = 0 Or p2 = 0 Then row = -1 Else row = CInt(Mid(nm, p2 + 1, Len(nm) - p2 - 1))
                If row < 0 Then row = maxrow + 1
                If row > maxrow Then maxrow = row
                If Not d.Exists(row) Then d.Add row, ""
                d(row) = d(row) & String(1, " ") & comp.Text
            End If
        Next
        For row = 0 To maxrow
            If d.Exists(row) Then cur = cur & d(row) & vbCrLf
        Next
        If cur = prev Or cur = "" Then Exit For
        log.Write cur
        prev = cur
        session.findById("wnd[0]").sendVKey 82   ' PgDn
        WScript.Sleep 600
    Next
End Sub

Sub Scan(term, regex, label)
    log.WriteLine "=============== SCAN: " & label & "  [" & term & "] regex=" & regex & " ==============="
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1000
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "RPR_ABAP_SOURCE_SCAN"
    session.findById("wnd[0]").sendVKey 8
    WScript.Sleep 1500
    With session
        .findById("wnd[0]/usr/txtSSTRING-LOW").Text = term
        .findById("wnd[0]/usr/chkP_REGEX").Selected = regex
        .findById("wnd[0]/usr/ctxtREPNAME-LOW").Text = "Z33SDR600"
        .findById("wnd[0]/usr/txtPLUSMINU").Text = "0"
        .findById("wnd[0]/usr/chkINCLU").Selected = True
        .findById("wnd[0]/usr/chkCOMMENT").Selected = True
        .findById("wnd[0]/usr/radRB_CODE").Select
        .findById("wnd[0]").sendVKey 8
    End With
    WScript.Sleep 4000
    log.WriteLine "  screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & " sbar=[" & session.findById("wnd[0]/sbar").Text & "]"
    On Error Resume Next
    Dim w1 : Set w1 = session.findById("wnd[1]")
    If Err.Number = 0 Then log.WriteLine "  popup: " & w1.Text : w1.sendVKey 0 : WScript.Sleep 800
    Err.Clear
    On Error GoTo 0
    If session.Info.ScreenNumber <> 1000 Then DumpList
End Sub

Scan "INCLUDE ", False, "?????? ???"
Scan "CALL FUNCTION|BAPI_|CALL TRANSACTION|SUBMIT ", True, "???/BAPI ???"
Scan "MESSAGE |TEXT-M|MSGV|BAPIRET", True, "?????"
Scan "WHEN '|CHANGE_MAT|CRE_DELIVERY|DEL_DELIVERY|GOODS|PGI|POST", True, "????? จจ???"
Scan "FROM Z|MODIFY Z|UPDATE Z|INSERT Z|INSERT INTO Z|DELETE Z|DELETE FROM Z|TABLES Z|TYPE Z", True, "Z ????? ????"

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "scan done"
