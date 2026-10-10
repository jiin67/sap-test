' test_hardcopy.vbs : check which HardCopy image formats this SAP GUI supports
'   READ ONLY. Writes sample captures into results\shots\
Option Explicit
Dim SapGuiAuto, application, connection, base, session, fso, log, n0, i

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
If connection Is Nothing Then
    WScript.Echo "GUARD: no KDQ connection"
    WScript.Quit 99
End If
Set base = connection.Children.ElementAt(0)

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\sap-test\results\shots") Then fso.CreateFolder "C:\sap-test\results\shots"
Set log = fso.CreateTextFile("C:\sap-test\logs\test_hardcopy.txt", True, True)

n0 = connection.Children.Count
If n0 > 1 Then
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

' go to a screen with content
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCRM052"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 2200
Err.Clear
On Error GoTo 0

Dim variants, v, path, res
' HardCopy(FileName, ImageType) - ImageType differs by GUI release; try the usual ones
For v = 0 To 3
    path = "C:\sap-test\results\shots\probe_type" & v & ".png"
    On Error Resume Next
    res = session.findById("wnd[0]").HardCopy(path, v)
    If Err.Number <> 0 Then
        log.WriteLine "HardCopy(path," & v & ") ERR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        log.WriteLine "HardCopy(path," & v & ") OK returned=[" & res & "] exists=" & fso.FileExists(path)
        If fso.FileExists(path) Then log.WriteLine "   size=" & fso.GetFile(path).Size
    End If
    On Error GoTo 0
Next

' single-argument form
path = "C:\sap-test\results\shots\probe_single.bmp"
On Error Resume Next
res = session.findById("wnd[0]").HardCopy(path)
If Err.Number <> 0 Then
    log.WriteLine "HardCopy(path) ERR " & Err.Number & " " & Err.Description
    Err.Clear
Else
    log.WriteLine "HardCopy(path) OK returned=[" & res & "] exists=" & fso.FileExists(path)
    If fso.FileExists(path) Then log.WriteLine "   size=" & fso.GetFile(path).Size
End If
On Error GoTo 0

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
log.WriteLine "done"
log.Close
session.findById("wnd[0]").close
WScript.Echo "test_hardcopy done"
