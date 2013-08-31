' This script installs Any2fb2 shell extensions:
' "Autoconvert to fb2" menu item for the files of types
' *.html; *.htm; *.txt; *.prt; *.rtf; *.doc; *.dot; *.wri; *.wk1; *.wk3; *.wk4; *.mcw
'
' Use "installShellExtensions.vbs -u" to remove all corresponding registry entries.
' 
' Script must be executed from the same folder where Any2fb2.vbs is located.
'
' Fj (fj.mail@gmail.com)

Dim Sh 'shell object
Dim cmdline 'open command line
Dim supportedExtensions
supportedExtensions = Split(".html .htm .txt .prt .rtf .doc .dot .wri .wk1 .wk3 .wk4 .mcw")
Dim uninstall 'true to remove keys

' *******
Function WriteApplicationKey(key)
	Sh.RegWrite key & "FriendlyAppName", "Autoconvert to fb2"
	
	Sh.RegWrite key & "shell\", "Autoconvert2fb2"
	Sh.RegWrite key & "shell\Autoconvert2fb2\", "&Autoconvert to fb2"
	Sh.RegWrite key & "shell\Autoconvert2fb2\FriendlyAppName", "Autoconvert to fb2"
	Sh.RegWrite key & "shell\Autoconvert2fb2\command\", cmdline, "REG_EXPAND_SZ"
	
	For Each ext in supportedExtensions 	
		Sh.RegWrite key & "SupportedTypes\" & ext, "", "REG_SZ"
	Next 
End Function

' *******
Function DeleteApplicationKey(key)
	Sh.RegDelete key & "shell\Autoconvert2fb2\Command\"
	Sh.RegDelete key & "shell\Autoconvert2fb2\"
	Sh.RegDelete key & "shell\"
	Sh.RegDelete key & "SupportedTypes\"
	Sh.RegDelete key
End Function

' *******
' *******
Function AssociateWithExtension(ext)
	Dim key

	On Error Resume Next
	filetype = Sh.RegRead ("HKEY_CLASSES_ROOT\" & ext & "\")
	On Error GoTo 0
	
	If filetype <> Empty Then
		key = "HKEY_CLASSES_ROOT\" & filetype & "\"
		Sh.RegWrite key & "shell\Autoconvert2fb2\", "&Autoconvert to fb2"
		Sh.RegWrite key & "shell\Autoconvert2fb2\Command\", cmdline, "REG_EXPAND_SZ"
	End If
End Function

' *******
Function UnAssociateWithExtension(ext)
	Dim key

	filetype = Sh.RegRead ("HKEY_CLASSES_ROOT\" & ext & "\")
	
	If filetype <> Empty Then
		key = "HKEY_CLASSES_ROOT\" & filetype & "\"
		Sh.RegDelete key & "shell\Autoconvert2fb2\Command\"
		Sh.RegDelete key & "shell\Autoconvert2fb2\"
	End If
End Function




' =============================
Set Sh = CreateObject("WScript.Shell")
cmdline = """%SystemRoot%\System32\WScript.exe"" """ & Sh.CurrentDirectory & "\any2fb2.vbs"" ""%1"""

' parse args

uninstall = False
For Each arg in WScript.Arguments
	If arg = "-u" Then uninstall = True
Next

If uninstall Then
	On Error Resume Next
	' base application key
	DeleteApplicationKey("HKEY_CLASSES_ROOT\Any2fb2.vbs\")
	' "open with" menu
	' DeleteApplicationKey("HKEY_CLASSES_ROOT\Applications\Any2fb2.vbs\")
	' file extensions
	For Each ext in supportedExtensions
		UnAssociateWithExtension(ext)
	Next
	On Error GoTo 0
	WSCript.Echo "Registry setup completed, ""Autoconvert to fb2"" option removed"
Else
	' base application key
	WriteApplicationKey("HKEY_CLASSES_ROOT\Any2fb2.vbs\")
	' "open with" menu
	' WriteApplicationKey("HKEY_CLASSES_ROOT\Applications\Any2fb2.vbs\")
	' file extensions
	For Each ext in supportedExtensions
		AssociateWithExtension(ext)
	Next
	WSCript.Echo "Registry setup completed, ""Autoconvert to fb2"" option installed, use ""installShellExtensions.vbs -u"" to remove"
End If

