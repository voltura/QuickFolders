@ECHO OFF
SETLOCAL

SET CSC="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
SET SRC=Program.cs
SET OUT=QuickFolders.exe
SET RES=QuickFolders.res
SET VB="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Microsoft.VisualBasic.dll"

"C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x86\rc.exe" QuickFolders.rc

IF EXIST %OUT% (
    ECHO Cleaning up existing %OUT%...
    DEL /F /Q %OUT%
)

ECHO Compiling %OUT%...

%CSC% /nologo /target:winexe /platform:x86 /optimize+ /debug- /nowarn:1591 /filealign:512 /win32res:%RES% /main:Program /out:%OUT% /reference:%VB% ^
/resource:Resources\1.png,QuickFolders.Resources.1.png ^
/resource:Resources\2.png,QuickFolders.Resources.2.png ^
/resource:Resources\3.png,QuickFolders.Resources.3.png ^
/resource:Resources\4.png,QuickFolders.Resources.4.png ^
/resource:Resources\5.png,QuickFolders.Resources.5.png ^
/resource:Resources\darkmode.png,QuickFolders.Resources.darkmode.png ^
/resource:Resources\exit.png,QuickFolders.Resources.exit.png ^
/resource:Resources\folder.png,QuickFolders.Resources.folder.png ^
/resource:Resources\lightmode.png,QuickFolders.Resources.lightmode.png ^
/resource:Resources\link.png,QuickFolders.Resources.link.png ^
/resource:Resources\more.png,QuickFolders.Resources.more.png ^
/resource:Resources\system.png,QuickFolders.Resources.system.png ^
/resource:Resources\theme.png,QuickFolders.Resources.theme.png ^
/resource:Resources\x.png,QuickFolders.Resources.x.png ^
/resource:Resources\bolt.png,QuickFolders.Resources.bolt.png ^
/resource:Resources\1_dark.png,QuickFolders.Resources.1_dark.png ^
/resource:Resources\2_dark.png,QuickFolders.Resources.2_dark.png ^
/resource:Resources\3_dark.png,QuickFolders.Resources.3_dark.png ^
/resource:Resources\4_dark.png,QuickFolders.Resources.4_dark.png ^
/resource:Resources\5_dark.png,QuickFolders.Resources.5_dark.png ^
/resource:Resources\darkmode_dark.png,QuickFolders.Resources.darkmode_dark.png ^
/resource:Resources\exit_dark.png,QuickFolders.Resources.exit_dark.png ^
/resource:Resources\folder_dark.png,QuickFolders.Resources.folder_dark.png ^
/resource:Resources\lightmode_dark.png,QuickFolders.Resources.lightmode_dark.png ^
/resource:Resources\link_dark.png,QuickFolders.Resources.link_dark.png ^
/resource:Resources\more_dark.png,QuickFolders.Resources.more_dark.png ^
/resource:Resources\system_dark.png,QuickFolders.Resources.system_dark.png ^
/resource:Resources\theme_dark.png,QuickFolders.Resources.theme_dark.png ^
/resource:Resources\x_dark.png,QuickFolders.Resources.x_dark.png ^
/resource:Resources\bolt_dark.png,QuickFolders.Resources.bolt_dark.png ^
%SRC%

IF NOT EXIST %OUT% (
    ECHO Build failed.
    ENDLOCAL
    EXIT /B 1
)

ECHO Build complete. Final file size:
FOR %%F IN (%OUT%) DO ECHO %%~zF bytes

ECHO Injecting application manifest to make application DPI aware.
"C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x86\mt.exe" -manifest app.manifest -outputresource:QuickFolders.exe;#1

ECHO Creating installer
"C:\Program Files (x86)\NSIS\makensis.exe" installer.nsi

START QuickFolders-Setup.exe

ENDLOCAL
