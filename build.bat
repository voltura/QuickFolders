@ECHO OFF
SETLOCAL

SET CSC="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
SET SRC=Program.cs
SET OUT=QuickFolders.exe
SET RES=explorer.res

ECHO [1/3] Cleaning up existing %OUT%...
IF EXIST %OUT% (
    DEL /F /Q %OUT%
)

ECHO [2/3] Compiling %OUT%...
%CSC% ^ 
  /nologo ^ 
  /target:winexe ^ 
  /platform:x86 ^ 
  /optimize+ ^ 
  /debug- ^ 
  /nowarn:1591 ^ 
  /filealign:512 ^ 
  /win32res:%RES% ^ 
  /main:P ^ 
  /out:%OUT% ^ 
  %SRC%

IF NOT EXIST %OUT% (
    ECHO Build failed.
    ENDLOCAL
    EXIT /B 1
)

ECHO [3/3] Build complete. Final file size:
FOR %%F IN (%OUT%) DO ECHO %%~zF bytes

ENDLOCAL
