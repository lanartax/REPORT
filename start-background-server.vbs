Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\Inventaris SJS\Documents\REPORT"
WshShell.Run """C:\Program Files\nodejs\node.exe"" local-server.js", 0, False

