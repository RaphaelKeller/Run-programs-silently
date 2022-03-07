Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c <Path_of_batch_or_file>"
oShell.Run strArgs, 0, false