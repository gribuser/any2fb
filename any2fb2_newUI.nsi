SetCompressor bzip2
; Default names
!define NAME "Any to FB2"
!define NSPNAME "Any2FB2"
!define VENDOR "GribUser"
!define VERSION "1.0"

!define MUI_PRODUCT "Any2FB2"
!define MUI_VERSION "1.0" ;Define your own software version here

; The name of the installer
;Name "${NAME}"
!include "MUI.nsh"

  !define MUI_LICENSEPAGE
  !define MUI_COMPONENTSPAGE
  !define MUI_DIRECTORYPAGE
  
  !define MUI_ABORTWARNING
  
  !define MUI_UNINSTALLER
  !define MUI_UNCONFIRMPAGE
  !insertmacro MUI_LANGUAGE English

; The file to write
OutFile "Any2FB2.exe"

; License text
;LicenseText "You must read the following license before installing:"
LicenseData "LICENSE.txt"

; The default installation directory
InstallDir "$PROGRAMFILES\${VENDOR}\${NAME}"
InstallDirRegKey HKLM "SOFTWARE\${VENDOR}\${NSPNAME}" "InstallDir"

; The text to prompt the user to enter a directory
;DirText "Please select a location to install ${NAME} (or use the default):"

; other settings
ShowInstDetails hide
ShowUninstDetails show

LangString DESC_SecMain ${LANG_ENGLISH} "REQUIRED to run Any2FB2.$\nRegisters an import plug-in for FBE and ActiveX server. You can use Any2fb2 from external scripts (JS, VBS, Perl…) and applications (FBE, Word…)"
LangString DESC_SecScripts ${LANG_ENGLISH} '(Will not work without dll) A set of scripts allowing you to use Any2fb2 directly. Use it as examples of ActiveX server script usage. Command-line and GUI VBS scripts are available'
LangString DESC_SecSorces ${LANG_ENGLISH} 'All required files to build Any2fb with Delphi. Any2FB is developed for Delphi 5, but any later versions of Delphi are supported. All required files (RXLib f.e.) except msxml_tlb.pas included.'
LangString DESC_SecShell ${LANG_ENGLISH} 'Adds "Convert to FB2" item to TXT|HTML files context menu in explorer.'
LangString DESC_SubSecScripts ${LANG_ENGLISH} 'Two Visial Basic samples, allowing you to use Any2FB2. Optionally you can add shell context menu to run one of this scripts on any TXT|HTML file from Explorer context menu'
LangString DESC_SecPythBatch ${LANG_ENGLISH} 'Really flexible python batch script from Alex Shabarshoff [shura@uc.ru].$\nAllows you to convert whole folders, list of files, collects erroneous documents, generates reports and so on.'

InstType "Recommended"
InstType "Integrated with shell"
InstType "Developer"

Function .onInit
	GetDLLVersion "$SYSDIR\vbscript.dll" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntCmpU 5 $R2 VBScriptIsOk1 VBScriptIsOk1 VBSCriptISTooOld
	VBScriptIsOk1:
	IntCmpU 5 $R3 VBScriptIsOk2 VBScriptIsOk2 VBSCriptISTooOld
	VBSCriptISTooOld:
		MessageBox MB_YESNO "Your WSH version is too old ($R2.$R3). You must have at least WHS v5.5. Please install the latest Windows Scripting Host and then install Any2FB$\n$\nShould installer now redirect you to Microsoft download page?$\n(choose your OS and language, then download and run installer, ~650kb)" IDNO no
		ExecShell "" "http://www.microsoft.com/downloads/results.aspx?freetext=wsh" SW_SHOWNORMAL
	no:
	MessageBox MB_OK|MB_ICONSTOP "Installation aborted!$\nPlease install WSH and retry."
	Quit
	VBScriptIsOk2:

	GetDLLVersion "$SYSDIR\msxml4.dll" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntCmpU 4 $R2 MSXMLIsOK1 MSXMLIsOK1 MSXMLIsTooOld
	MSXMLIsOK1:
	IntCmpU 10 $R3 MSXMLIsOK MSXMLIsOK MSXMLIsTooOld
	MSXMLIsTooOld:
		MessageBox MB_YESNO "Your MSXML version is too old ($R2.$R3). You must have at least 4.10. Please install the latest MSXML and then install FB2Any$\n$\nShould installer now start downloading 'MSXML 4.0 sp2' (~5.0Mb) for you? " IDNO no1
		ExecShell "" "http://www.gribuser.ru/xml/fictionbook/2.0/software/msxml_4.0_sp2.msi" SW_SHOWNORMAL
	no1:
	MessageBox MB_OK|MB_ICONSTOP "Installation aborted!$\nPlease install MSXML 4.0 and retry."
	Quit
	MSXMLIsOK:
FunctionEnd

Section ""
  ; prepare install env
	SetShellVarContext All
  SetOutPath $INSTDIR
  CreateDirectory "$SMPROGRAMS\${NAME}"
SectionEnd


Section "!Any2FB2 main dll" SecMain

	SetShellVarContext All
  IfFileExists "$INSTDIR\any_2_fb2.dll" 0 nodll
    UnRegDll "$INSTDIR\any_2_fb2.dll"
  nodll:
  File "any_2_fb2.dll"
  ; register application
  RegDll "$INSTDIR\any_2_fb2.dll"
	; Header Detection boosted preset
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Remove External Links" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Remove External Links" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Remove ALL images" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Remove off-site images" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Preserve dinamic images" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "No enmty lines" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Leave quotes as is" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Skip footnotes" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Skip _italic_" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Skip epigraphs" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "No paragraph restore" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "No description" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Allow 500 errors" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Skip poems" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Preserv forms" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "No encoding detection" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Header detect method" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "NO headers detection" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Ignore spaces" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Leave dashes" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Links download level" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Folow external links" 0
	WriteRegDWORD HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Folow links" 0
	WriteRegStr HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+" "Headers detect regexp" "(^\s*Глава\s.{1,25}$$|^\s*Часть\s.{1,25}$$|^\s*\*{3,}\s*$$|^\s*Книга\s.{5,25}$$|^\s*От автора\s*$$|^\s*Пролог\s*$$|^\s*Эпилог\s*$$|^\s*Chapter\s.{1,25}$$|^\s*Epilog\s*$$|^\s*Prologue\s*$$)"
	WriteRegStr HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+\onload" "0" `<p[^>]+align=(["']?)center\1[^>]+>`
	WriteRegStr HKCU "Software\Grib Soft\Any to FB2\1.0\presets\Headers+\onload" "0_" "<h1>"
	SectionIn 1
	SectionIn 2
SectionEnd

SubSection "Scripts" SubSecScripts
	Section "VBS samples" SecScripts
		SetShellVarContext All
	  File "any2fb2.vbs"
	  File "any2fb2_gui.vbs"
		File "installShellExtensions.vbs"
		SetOutPath "$SYSDIR"
		File "wscript.exe.manifest"
	  SetOutPath $INSTDIR
	  CreateShortCut "$SMPROGRAMS\${NAME}\Any to FB2 GUI.lnk" "$INSTDIR\any2fb2_gui.vbs" "" "$INSTDIR\any_2_fb2.dll"
	  CreateShortCut "$SMPROGRAMS\${NAME}\Any to FB2 command line.lnk" "$INSTDIR\any2fb2.vbs" "" "shell32.dll" "2"
		SectionIn 1
		SectionIn 2
		SectionIn 3
	SectionEnd
	Section "Shell context menu" SecShell
		WriteRegStr HKCR "txtfile\shell\ToFB2" "" "&Convert to FB2..."
		WriteRegStr HKCR "txtfile\shell\ToFB2\command" "" 'wscript.exe "$INSTDIR\any2fb2_gui.vbs" "%L"'
		WriteRegStr HKCR "htmlfile\shell\ToFB2" "" "&Convert to FB2..."
		WriteRegStr HKCR "htmlfile\shell\ToFB2\command" "" 'wscript.exe "$INSTDIR\any2fb2_gui.vbs" "%L"'
		SectionIn 2
	SectionEnd
SubSectionEnd

Section "Python Batch script" SecPythBatch
  SetOutPath $INSTDIR
	file "Any2FBatch.py"
  CreateShortCut "$SMPROGRAMS\${NAME}\Any2FBbatch.py.lnk" "$INSTDIR\Any2FBatch.py" "" ""
SectionEnd

Section "Sources (Delphi 5)" SecSorce
	SetShellVarContext All
  CreateDirectory "$INSTDIR\src"
  CreateDirectory "$INSTDIR\src"
	SetOutPath "$INSTDIR\src"

	File "src\maxmin.pas"
	File "src\icolist.pas"
	File "src\dispatchforany2fb.pas"
	File "src\clipicon.pas"
	File "src\any_2_fb_dialog.pas"
	File "src\any_2_fb2_tlb.pas"
	File "src\anifile.pas"
	File "src\any_2_fb2.dpr"
	File "src\reeditor.dfm"
	File "src\presetslist.dfm"
	File "src\any_2_fb_dialog.dfm"
	File "src\rxgconst.r32"
	File "src\rxconst.r32"
	File "src\rxcconst.r32"
	File "src\rx.inc"
	File "src\any_2_fb2.tlb"
	File "src\icon.res"
	File "src\vclutils.pas"
	File "src\unit1.pas"
	File "src\tx2fb.pas"
	File "src\rxgraph.pas"
	File "src\rxgif.pas"
	File "src\rxgconst.pas"
	File "src\rxconst.pas"
	File "src\rxcconst.pas"
	File "src\reeditor.pas"
	File "src\presetslist.pas"
	File "src\pngzlib.pas"
	File "src\pnglang.pas"
	File "src\pngimage.pas"
	File "src\msregexpr.pas"

	CreateDirectory "$INSTDIR\src\obj"
	SetOutPath "$INSTDIR\src\obj"
	File "src\obj\inftrees.obj"
	File "src\obj\inflate.obj"
	File "src\obj\inffast.obj"
	File "src\obj\infcodes.obj"
	File "src\obj\infblock.obj"
	File "src\obj\deflate.obj"
	File "src\obj\adler32.obj"
	File "src\obj\trees.obj"
	File "src\obj\infutil.obj"

  CreateShortCut "$SMPROGRAMS\${NAME}\Any to FB2 Project.lnk" "$INSTDIR\src\any_2_fb2.dpr"
	SectionIn 3
SectionEnd

Section "" ;; finalize install 3 0
  SetOutPath $INSTDIR
	SetShellVarContext All
  ; Uninstall shortcut
  CreateShortCut "$SMPROGRAMS\${NAME}\Uninstall.lnk" "$INSTDIR\uninstall.exe" "" "$INSTDIR\uninstall.exe" 0
  ; Write the installation path into the registry
  WriteRegStr HKLM "SOFTWARE\${VENDOR}\${NSPNAME}" "InstallDir" "$INSTDIR"
  ; Uninstall info
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}" "DisplayName" "${VENDOR} ${NAME} ${VERSION} (remove only)"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  ; uninstall program
  WriteUninstaller "uninstall.exe"

	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}" "NoModify" 1
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}" "NoRepair" 1
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}" "DisplayVersion" '${MUI_VERSION}'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}" "URLInfoAbout" 'http://www.gribuser.ru'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}" "Publisher" '${VENDOR}'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}" "URLUpdateInfo" 'http://www.gribuser.ru/xml/fictionbook/2.0/software/'
SectionEnd

; Uninstall support
;UninstallText "This will uninstall ${VENDOR} ${NAME}. Hit Uninstall to continue."


Section "Uninstall"
  ; remove plugin
	SetShellVarContext All
  UnRegDll "$INSTDIR\any_2_fb2.dll"
  Delete "$INSTDIR\any_2_fb2.dll"
  Delete "$INSTDIR\any2fb2.vbs"
  Delete "$INSTDIR\any2fb2_gui.vbs"
	Delete "$INSTDIR\installShellExtensions.vbs"
	Delete "$INSTDIR\Any2FBatch.py"
	Delete "$SYSDIR\wscript.exe.manifest"
	RMDir /r "$INSTDIR\src"

  ; MUST REMOVE UNINSTALLER, too
  Delete "$INSTDIR\uninstall.exe"

  ; remove directories used.
  RMDir /r "$SMPROGRAMS\${NAME}"
  RMDir "$INSTDIR"
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${VENDOR} ${NSPNAME}"
	DeleteRegKey HKCR "htmlfile\shell\ToFB2"
	DeleteRegKey HKCR "txtfile\shell\ToFB2"
SectionEnd

!insertmacro MUI_FUNCTIONS_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecMain} $(DESC_SecMain)
  !insertmacro MUI_DESCRIPTION_TEXT ${SecScripts} $(DESC_SecScripts)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecSorce} $(DESC_SecSorces)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecShell} $(DESC_SecShell)
	!insertmacro MUI_DESCRIPTION_TEXT ${SubSecScripts} $(DESC_SubSecScripts)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecPythBatch} $(DESC_SecPythBatch)
!insertmacro MUI_FUNCTIONS_DESCRIPTION_END
