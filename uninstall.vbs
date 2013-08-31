Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
set WshShell = WScript.CreateObject("WScript.Shell")
strPrograms = WshShell.SpecialFolders("AllUsersPrograms")

on error resume next

Set RegExprObj = New RegExp
RegExprObj.Pattern = "(.*)\\[^\\]*"
path = RegExprObj.Replace(WScript.ScriptFullName,"$1")

fso.DeleteFile(strPrograms & "\Any 2 FB2\Any 2 fb2 GUI.lnk")
fso.DeleteFile(strPrograms & "\Any 2 FB2\Any 2 fb2 Command-line.lnk")
fso.DeleteFile(strPrograms & "\Any 2 FB2\Uninstall.lnk")
if FolderEmpty(strPrograms & "\Any 2 FB2") then
	fso.DeleteFolder(strPrograms & "\Any 2 FB2")
end if

WshShell.run "REGSVR32.EXE /u /s """ & path & "\any_2_fb2.dll"""
fso.DeleteFile(path & "\any2fb2.vbs")
fso.DeleteFile(path & "\any2fb2_gui.vbs")
fso.DeleteFile(path & "\any_2_fb2.dll")
fso.DeleteFile(WScript.ScriptFullName)
if FolderEmpty(path) then
	fso.DeleteFolder(path)
end if

MsgBox "Uninstall complete!"




function FolderEmpty(path)
	if fso.GetFolder(path).Files.count > 0 then
		FolderEmpty = false
	else
		FolderEmpty = true
	end if
end function