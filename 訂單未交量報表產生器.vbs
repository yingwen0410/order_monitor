Set objShell = CreateObject("WScript.Shell")
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
pythonPath = scriptDir & "\system\python_portable\pythonw.exe"
mainPath = scriptDir & "\system\main.py"
' 1 = SW_SHOWNORMAL (show window, so Tkinter GUI doesn't inherit SW_HIDE), False = don't wait
objShell.Run """" & pythonPath & """ """ & mainPath & """", 1, False
