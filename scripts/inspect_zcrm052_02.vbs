' inspect_zcrm052_02.vbs : program attributes + source check for SUBS_TYPE condition
' READ ONLY. SE38 display mode only, never change mode.
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children.ElementAt(0)
Set base = connection.Children.ElementAt(0)
If base.Info.SystemName <> "KDQ" And base.Info.SystemName <> "KDD" Then
    WScript.Echo "GUARD: system " & base.Info.SystemName & " not allowed"
    WScript.Quit 99
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("C:\sap-test\logs\inspect_zcrm052_02.txt", True, True)
log.WriteLine "system=" & base.Info.SystemName & " client=" & base.Info.Client & " user=" & base.Info.User

n0 = connection.Children.Count
' Always work in a session we opened ourselves. Re-using a leftover session can
' block for a long time when that session still holds a huge ALV result.
If False Then
    Set session = connection.Children.ElementAt(n0 - 1)
Else
    base.createSession
    For i = 1 To 40
        WScript.Sleep 500
        If connection.Children.Count > n0 Then Exit For
    Next
    Set session = connection.Children.ElementAt(connection.Children.Count - 1)
    WScript.Sleep 1500
End If
log.WriteLine "work session=" & session.Id

Sub Status()
    log.WriteLine "  prog/screen=" & session.Info.Program & "/" & session.Info.ScreenNumber & "  tcode=" & session.Info.Transaction
    log.WriteLine "  title=[" & session.findById("wnd[0]").Text & "]"
    log.WriteLine "  sbar =[" & session.findById("wnd[0]/sbar").Text & "] type=" & session.findById("wnd[0]/sbar").MessageType
End Sub

Sub DumpFlat(root)
    Dim c, k, comp, line
    On Error Resume Next
    k = root.Children.Count
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    For c = 0 To k - 1
        Set comp = root.Children.ElementAt(c)
        line = comp.Type & " | " & comp.Id & " | text=[" & comp.Text & "]"
        log.WriteLine line
        DumpFlat comp
    Next
    Err.Clear
End Sub

' ---- attributes ----
Sub ProgAttr(pname)
    log.WriteLine "=============== SE38 ATTRIBUTES " & pname & " ==============="
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1300
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = pname
    session.findById("wnd[0]/usr/radRS38M-FUNC_HEAD").Select
    session.findById("wnd[0]/usr/btnSHOP").press
    WScript.Sleep 2200
    If Err.Number <> 0 Then log.WriteLine "  ERR " & Err.Number & " " & Err.Description : Err.Clear
    On Error GoTo 0
    Status
    DumpFlat session.findById("wnd[0]/usr")
End Sub

ProgAttr "ZCRMR052"
ProgAttr "ZCRMR052S"

' ---- source of ZCRMR052S : read editor text, grep SUBS_TYPE ----
Sub DumpSource(pname, outfile)
    Dim ed, txt, arr, j, f, hit
    log.WriteLine "=============== SE38 SOURCE " & pname & " ==============="
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE38"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 1300
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = pname
    session.findById("wnd[0]/usr/radRS38M-FUNC_EDIT").Select
    session.findById("wnd[0]/usr/btnSHOP").press
    WScript.Sleep 3000
    If Err.Number <> 0 Then log.WriteLine "  ERR open " & Err.Number & " " & Err.Description : Err.Clear
    On Error GoTo 0
    Status

    On Error Resume Next
    Set ed = session.findById("wnd[0]/usr/cntlEDITOR_CONTAINER/shellcont/shell")
    If Err.Number <> 0 Then
        Err.Clear
        Set ed = session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell")
    End If
    If Err.Number <> 0 Then
        Err.Clear
        log.WriteLine "  editor shell not found, dumping usr tree"
        DumpFlat session.findById("wnd[0]/usr")
        Exit Sub
    End If
    txt = ed.Text
    If Err.Number <> 0 Then
        log.WriteLine "  ERR read editor text " & Err.Number & " " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Set f = fso.CreateTextFile(outfile, True, True)
    f.Write txt
    f.Close
    log.WriteLine "  source saved to " & outfile & "  chars=" & Len(txt)

    arr = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    log.WriteLine "  total lines=" & (UBound(arr) + 1)
    hit = 0
    For j = 0 To UBound(arr)
        If InStr(UCase(arr(j)), "SUBS_TYPE") > 0 Then
            log.WriteLine "  L" & (j + 1) & ": " & arr(j)
            hit = hit + 1
        End If
    Next
    log.WriteLine "  SUBS_TYPE hits=" & hit
End Sub

DumpSource "ZCRMR052S", "C:\sap-test\logs\src_ZCRMR052S.txt"

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 900
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "inspect_zcrm052_02 done"
