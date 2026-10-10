' probe_kdd.vbs : READ-ONLY. list connections / sessions and pick KDD.
Option Explicit
Dim SapGuiAuto, application, i, j, con, ses
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
For i = 0 To application.Children.Count - 1
    Set con = application.Children.ElementAt(i)
    For j = 0 To con.Children.Count - 1
        Set ses = con.Children.ElementAt(j)
        WScript.Echo "con[" & i & "] ses[" & j & "] sys=" & ses.Info.SystemName & " client=" & ses.Info.Client & " user=" & ses.Info.User & " lang=" & ses.Info.Language & " tcode=" & ses.Info.Transaction & " prog=" & ses.Info.Program & " busy=" & ses.Busy
    Next
Next
