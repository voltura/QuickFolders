@ECHO OFF
SETLOCAL

SET CSC="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
SET SRC=Program.cs
SET OUT=QuickFolders.exe
SET RES=QuickFolders.res
SET VB="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Microsoft.VisualBasic.dll"

"C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x86\rc.exe" QuickFolders.rc

IF EXIST %OUT% (
    ECHO Cleaning up existing %OUT%...
    DEL /F /Q %OUT%
)

ECHO Compiling %OUT%...

%CSC% /nologo /target:winexe /platform:x86 /optimize+ /debug- /nowarn:1591 /filealign:512 /win32res:%RES% /main:P /out:%OUT% /reference:%VB% ^
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
%SRC%

IF NOT EXIST %OUT% (
    ECHO Build failed.
    ENDLOCAL
    EXIT /B 1
)

ECHO Build complete. Final file size:
FOR %%F IN (%OUT%) DO ECHO %%~zF bytes

"C:\Program Files (x86)\NSIS\makensis.exe" installer.nsi

ENDLOCAL
