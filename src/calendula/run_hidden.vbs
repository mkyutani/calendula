Set shell = CreateObject("WScript.Shell")
shell.Run "powershell.exe -ExecutionPolicy Bypass -NonInteractive -File """ & WScript.Arguments(0) & """", 0, True
