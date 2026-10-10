Option Explicit
Dim SapGuiAuto, app, i, ok
ok = False
For i = 1 To 10
    On Error Resume Next
    Set SapGuiAuto = Nothing
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number = 0 Then
        Set app = SapGuiAuto.GetScriptingEngine
        If Err.Number = 0 Then
            WScript.Echo "attempt " & i & " : OK  connections=" & app.Children.Count
            ok = True
            Exit For
        End If
    End If
    WScript.Echo "attempt " & i & " : " & Err.Number & " " & Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 3000
Next
If ok Then WScript.Quit 0
WScript.Quit 1
