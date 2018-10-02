Dim objNet
Dim wshShell
Dim wshProcess

Dim domain, username, dc

Set objNet = CreateObject("WScript.NetWork") 
Set wshShell = CreateObject("Wscript.Shell")
Set wshProcess = wshShell.Environment("Process")

Domain = objNet.UserDomain
UserName = objNet.UserName
DC = wshProcess("LogonServer")

Set UserObj = GetObject("WinNT://" & domain & "/" & username)

loginScript = DC & "\netlogon\" & UserObj.LoginScript

On Error Resume Next
wshShell.Run(loginscript)

Set UserObj = Nothing
Set wshProcess = Nothing
Set wshShell = Nothing
Set objNet = Nothing
