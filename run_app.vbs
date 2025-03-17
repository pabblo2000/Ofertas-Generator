Set WshShell = CreateObject("WScript.Shell")
' Ejecuta run_app.ps1 de forma oculta (0 para ventana oculta)
WshShell.Run "powershell -ExecutionPolicy Bypass -File """ & CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\run_app.ps1""", 0
Set WshShell = Nothing
