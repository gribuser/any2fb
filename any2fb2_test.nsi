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
OutFile "Any2FB2_test.exe"

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

Section "!Any2FB2 main dll" SecMain
	SetShellVarContext All
  SetOutPath $INSTDIR
  File "any_2_fb2.dll"
  ; register application

  Push $R0
  ReadRegStr $R0 HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion" CurrentVersion
  StrCmp $R0 "" 0 lbl_winnt
		Exec "$\"$SYSDIR\regsvr32.exe$\" /s $\"$INSTDIR\any_2_fb2.dll$\""
  Goto lbl_done
  lbl_winnt:
  RegDll "$INSTDIR\any_2_fb2.dll"
  lbl_done:
	Pop $R0
SectionEnd

