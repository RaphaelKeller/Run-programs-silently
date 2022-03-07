Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c <Path_of_batch_or_exe>"
oShell.Run strArgs, 0, false