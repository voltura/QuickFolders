SetCompressor /SOLID lzma
SetCompressorDictSize 128

!include LogicLib.nsh
!include nsDialogs.nsh
!insertmacro NSD_FUNCTION_INIFILE

!define APPNAME "QuickFolders"
!define COMPANY "Voltura AB"
!define VERSION "1.0.0.2"
!define INSTALLDIR "$LOCALAPPDATA\${APPNAME}"

OutFile "${APPNAME}-Setup.exe"
Icon "icon.ico"
UninstallIcon "icon.ico"
InstallDir "${INSTALLDIR}"
RequestExecutionLevel user

Name "${APPNAME} ${VERSION}"
InstallDirRegKey HKCU "Software\QuickFolders" "Install_Dir"

Page custom nsDialogsIO UpdateINIState
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

Function KillRunningApp
    nsExec::Exec 'tasklist /FI "IMAGENAME eq QuickFolders.exe" | find /I "QuickFolders.exe"'
    Pop $0
    StrCmp $0 "" done
    nsExec::Exec 'taskkill /F /IM QuickFolders.exe'
	done:
FunctionEnd

Function un.KillRunningApp
    nsExec::Exec 'tasklist /FI "IMAGENAME eq QuickFolders.exe" | find /I "QuickFolders.exe"'
    Pop $0
    StrCmp $0 "" done
    nsExec::Exec 'taskkill /F /IM QuickFolders.exe'
	done:
FunctionEnd

Function nsDialogsIO
	InitPluginsDir
	File /oname=$PLUGINSDIR\io.ini ".\installer.ini"
		StrCpy $0 $PLUGINSDIR\io.ini
	Call CreateDialogFromINI
FunctionEnd

Section "Install"
	Call KillRunningApp
	SetOutPath "$INSTDIR"
    File "QuickFolders.exe"
    CreateShortcut "$SMPROGRAMS\${APPNAME}.lnk" "$INSTDIR\QuickFolders.exe"
    WriteUninstaller "$INSTDIR\Uninstall.exe"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME} ${VERSION}"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "InstallLocation" "$INSTDIR"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" "$INSTDIR\QuickFolders.exe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" "${VERSION}"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "${COMPANY}"
    ReadINIStr $0 "$PLUGINSDIR\io.ini" "Field 1" "State"
    ReadINIStr $1 "$PLUGINSDIR\io.ini" "Field 2" "State"
    StrCmp $0 "1" startWithWindows
    Goto done
    startWithWindows:
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "QuickFolders" "$INSTDIR\QuickFolders.exe"
    done:
    StrCmp $1 "1" launchApp
    Goto done2
    launchApp:
    Exec "$INSTDIR\QuickFolders.exe"
    done2:
SectionEnd

Section "Uninstall"
    Call un.KillRunningApp
    Delete "$INSTDIR\QuickFolders.exe"
    Delete "$INSTDIR\Uninstall.exe"
    Delete "$SMPROGRAMS\${APPNAME}.lnk"
    DeleteRegKey HKCU "Software\QuickFolders"
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Run\QuickFolders"
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
	RMDir "$INSTDIR"
SectionEnd
