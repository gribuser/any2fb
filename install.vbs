Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
set WshShell = WScript.CreateObject("WScript.Shell")
strPrograms = WshShell.SpecialFolders("AllUsersPrograms")
on error resume next
fso.CreateFolder(strPrograms & "\Any 2 FB2")

Set RegExprObj = New RegExp
RegExprObj.Pattern = "(.*\\)[^\\]*"
path = RegExprObj.Replace(WScript.ScriptFullName,"$1")

set oShellLink = WshShell.CreateShortcut(strPrograms & "\Any 2 FB2\Any 2 fb2 GUI.lnk")
oShellLink.TargetPath = path & "any2fb2_gui.vbs"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = path & "any_2_fb2.dll, 0"
oShellLink.Description = "Graphic user interface for ""Any 2 fb"" library"
oShellLink.WorkingDirectory = path
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(strPrograms & "\Any 2 FB2\Any 2 fb2 Command-line.lnk")
oShellLink.TargetPath = path & "any2fb2.vbs"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = path & "shell32.dll, 2"
oShellLink.Description = "Command line interface for ""Any 2 fb"" library"
oShellLink.WorkingDirectory = path
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(strPrograms & "\Any 2 FB2\Uninstall.lnk")
oShellLink.TargetPath = path & "uninstall.vbs"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = path & "shell32.dll, 31"
oShellLink.Description = "Uninstall Any 2 FB2 converter"
oShellLink.WorkingDirectory = path
oShellLink.Save

WshShell.run "REGSVR32.EXE /s """ & path & "any_2_fb2.dll"""

fso.DeleteFile(WScript.ScriptFullName)