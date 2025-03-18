Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists("config.py") Then
    windowStyle = 0
Else
    windowStyle = 1
End If

Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell -ExecutionPolicy Bypass -NoExit -File """ & fso.GetParentFolderName(WScript.ScriptFullName) & "\run.ps1""", windowStyle
Set WshShell = Nothing
Set fso = Nothing
WScript.Quit
