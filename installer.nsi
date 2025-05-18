ManifestDPIAware true
ManifestSupportedOS Win10
SetCompressor /SOLID lzma
SetCompressorDictSize 128

!include LogicLib.nsh
!include nsDialogs.nsh
!include WinMessages.nsh

!define APPNAME "QuickFolders"
!define COMPANY "Voltura AB"
!define VERSION "1.0.1.3"
!define INSTALLDIR "$LOCALAPPDATA\${APPNAME}"
!define APPDATADIR "$APPDATA\${APPNAME}"

Var /GLOBAL StartWithWindowsState
Var /GLOBAL checkbox
Var /GLOBAL uiFont

OutFile "${APPNAME}-Setup.exe"
Icon "icon.ico"
UninstallIcon "icon.ico"
AutoCloseWindow true
InstallDir "${INSTALLDIR}"
RequestExecutionLevel user

Name "${APPNAME} ${VERSION}"
Caption "${APPNAME} ${VERSION} Setup - ${COMPANY}"

Page custom StartDialog StartDialogLeave
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

Function StartDialog
    nsDialogs::Create 1018
    Pop $0

    CreateFont $uiFont "Segoe UI" 9
    SendMessage $0 ${WM_SETFONT} $uiFont 1

    ${NSD_CreateLabel} 10u 5u 100% 10u "QuickFolders will be installed on your computer."
    Pop $1
    SendMessage $1 ${WM_SETFONT} $uiFont 1

    ${NSD_CreateLabel} 10u 25u 100% 10u "You can choose if QuickFolders should start automatically"
    Pop $1
    SendMessage $1 ${WM_SETFONT} $uiFont 1

    ${NSD_CreateLabel} 10u 35u 100% 10u "with Windows when you log in."
    Pop $1
    SendMessage $1 ${WM_SETFONT} $uiFont 1

    ${NSD_CreateCheckbox} 10u 55u 100% 8u "Start with Windows"
    Pop $checkbox
    SendMessage $checkbox ${WM_SETFONT} $uiFont 1
    ${NSD_Check} $checkbox

    ${NSD_CreateLabel} 10u 95u 100% 10u "Click Install to continue."
    Pop $1
    SendMessage $1 ${WM_SETFONT} $uiFont 1

    nsDialogs::Show
FunctionEnd

Function StartDialogLeave
    ${NSD_GetState} $checkbox $StartWithWindowsState
FunctionEnd

Function .onInstSuccess
    MessageBox MB_OK "QuickFolders was installed successfully! Check the system tray!"
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
    StrCmp $StartWithWindowsState 1 startWithWindows
    Goto done
    startWithWindows:
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "QuickFolders" "$INSTDIR\QuickFolders.exe"
    done:
    Exec "$INSTDIR\QuickFolders.exe"
SectionEnd

Section "Uninstall"
    Call un.KillRunningApp
    Delete "$INSTDIR\QuickFolders.exe"
    Delete "$INSTDIR\Uninstall.exe"
    Delete "$SMPROGRAMS\${APPNAME}.lnk"
    DeleteRegKey HKCU "Software\QuickFolders"
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Run\QuickFolders"
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
    MessageBox MB_YESNO|MB_ICONQUESTION "Do you also want to remove QuickFolders settings?" IDNO skipRemove
    Delete "${APPDATADIR}\QuickFolders.config"
    RMDir "${APPDATADIR}"
    skipRemove:
    RMDir "$INSTDIR"
SectionEnd
