@ECHO OFF
TITLE QuickFolders Build Script
SETLOCAL ENABLEDELAYEDEXPANSION

:: ─────────────────────────────────────────────────────────────────────────
::   Define variables
:: ─────────────────────────────────────────────────────────────────────────

SET "SCRIPT_DIR=%~dp0"
SET "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
SET "MT=C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x86\mt.exe"
SET "SIGN_TOOL=C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x86\signtool.exe"
SET "MAKE_NSIS=C:\Program Files (x86)\NSIS\makensis.exe"
SET "SETUP_EXE=QuickFolders-Setup.exe"
SET "RC=C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x86\rc.exe"
SET "CSC=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
SET OUT=QuickFolders.exe
SET RES=QuickFolders.res
SET SHA_OUT=QuickFolders.sha512
SET "VB=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Microsoft.VisualBasic.dll"
SET UPDATE_VER=%1

:: ─────────────────────────────────────────────────────────────────────────
::   Collect .CS files to be compiled
:: ─────────────────────────────────────────────────────────────────────────

SET "SRC="
SET FIRST=1

FOR %%F IN (*.cs) DO (
    IF /I NOT "%%~nxF"=="AssemblyInfo.cs" (
        IF !FIRST! EQU 1 (
            SET SRC="%%~nxF"
            SET FIRST=0
        ) ELSE (
            SET SRC=!SRC! ^
"%%~nxF"
        )
    )
)

:: ─────────────────────────────────────────────────────────────────────────
::   Collect resource PNG images to be included in EXE
:: ─────────────────────────────────────────────────────────────────────────

SET "RESOURCES="
FOR %%F IN (Resources\*.png) DO (
    SET "FILE=%%~nxF"
    SET RESOURCES=!RESOURCES! ^
/resource:"Resources\!FILE!","QuickFolders.Resources.!FILE!"
)

:: ─────────────────────────────────────────────────────────────────────────
::   Tool validation
:: ─────────────────────────────────────────────────────────────────────────

IF NOT EXIST "%MT%" (
	CALL :ERROR_MESSAGE_EXIT "mt.exe not found at %MT%." 11
)

IF NOT EXIST "%MAKE_NSIS%" (
	CALL :ERROR_MESSAGE_EXIT "makensis.exe not found at %MAKE_NSIS%." 12
)

IF NOT EXIST "%RC%" (
	CALL :ERROR_MESSAGE_EXIT "rc.exe not found at %RC%." 13
)

IF NOT EXIST "%CSC%" (
	CALL :ERROR_MESSAGE_EXIT "csc.exe not found at %CSC%." 14
)

IF NOT EXIST "%VB%" (
	CALL :ERROR_MESSAGE_EXIT "Microsoft.VisualBasic.dll not found at %VB%." 15
)

:: ─────────────────────────────────────────────────────────────────────────
::   Update version
:: ─────────────────────────────────────────────────────────────────────────
IF "%UPDATE_VER%" NEQ "1" GOTO :SKIP_VERSION_UPDATE
CALL :DISP_MSG "Getting current version from AssemblyInfo.cs..." 0

TYPE "%SCRIPT_DIR%\AssemblyInfo.cs"|FINDSTR AssemblyFileVersion >"%SCRIPT_DIR%\VERSION_REPLACE.TXT"
SET /P AssemblyFileVersion=<"%SCRIPT_DIR%\VERSION_REPLACE.TXT"
DEL /F /Q "%SCRIPT_DIR%\VERSION_REPLACE.TXT"
SET "CurrentAssemblyFileVersion=%AssemblyFileVersion:~32,-3%"

FOR /F "tokens=1,2,3,4 delims=." %%G IN ("%CurrentAssemblyFileVersion%") DO (
    SET /A MAJOR=%%G
    SET /A MINOR=%%H
    SET /A BUILD=%%I
    SET /A REVISION=%%J
)

IF %REVISION% GEQ 9 (
    SET /A REVISION=0
    IF %BUILD% GEQ 9 (
        SET /A BUILD=0
        IF %MINOR% GEQ 9 (
            SET /A MINOR=0
            SET /A MAJOR+=1
        ) ELSE (
            SET /A MINOR+=1
        )
    ) ELSE (
        SET /A BUILD+=1
    )
) ELSE (
    SET /A REVISION+=1
)

SET NewVersion=%MAJOR%.%MINOR%.%BUILD%.%REVISION%
SET NewVersionComma=%MAJOR%,%MINOR%,%BUILD%,%REVISION%

CALL :DISP_MSG "Updating version to %NewVersion%" 0

:: ─────────────────────────────────────────────────────────────
::   Replace version using PowerShell (encoding-aware)
:: ─────────────────────────────────────────────────────────────

:: AssemblyInfo.cs — already saved as UTF-8 with BOM
powershell -Command "(Get-Content -Raw '%SCRIPT_DIR%\AssemblyInfo.cs') -replace '%CurrentAssemblyFileVersion%', '%NewVersion%' | Out-File -Encoding utf8 -NoNewline '%SCRIPT_DIR%\AssemblyInfo.cs'"

:: Program.cs
powershell -Command "(Get-Content -Raw '%SCRIPT_DIR%\Program.cs') -replace '%CurrentAssemblyFileVersion%', '%NewVersion%' | Out-File -Encoding utf8 -NoNewline '%SCRIPT_DIR%\Program.cs'"

:: installer.nsi
powershell -Command "(Get-Content -Raw '%SCRIPT_DIR%\installer.nsi') -replace '%CurrentAssemblyFileVersion%', '%NewVersion%' | Out-File -Encoding utf8 -NoNewline '%SCRIPT_DIR%\installer.nsi'"

:: app.manifest
powershell -Command "(Get-Content -Raw '%SCRIPT_DIR%\app.manifest') -replace '%CurrentAssemblyFileVersion%', '%NewVersion%' | Out-File -Encoding utf8 -NoNewline '%SCRIPT_DIR%\app.manifest'"

:: QuickFolders.rc - dot version (UTF-16 LE BOM)
powershell -Command ^
  "$f='%SCRIPT_DIR%\QuickFolders.rc'; $c=Get-Content -Raw -Encoding Unicode $f; $c2=$c -replace '\"%CurrentAssemblyFileVersion%\"','\"%NewVersion%\"'; $sw=[System.IO.StreamWriter]::new($f,$false,[System.Text.Encoding]::Unicode); $sw.Write($c2); $sw.Close()"

:: QuickFolders.rc - comma version (UTF-16 LE BOM)
powershell -Command ^
  "$f='%SCRIPT_DIR%\QuickFolders.rc'; $v='%CurrentAssemblyFileVersion%'.Replace('.', ','); $n='%NewVersionComma%'; $c=Get-Content -Raw -Encoding Unicode $f; $c2=$c -replace $v,$n; $sw=[System.IO.StreamWriter]::new($f,$false,[System.Text.Encoding]::Unicode); $sw.Write($c2); $sw.Close()"

:SKIP_VERSION_UPDATE

:: ─────────────────────────────────────────────────────────────────────────
::   Create resource file
:: ─────────────────────────────────────────────────────────────────────────

CALL :DISP_MSG "Creating resource file..." 0

"%RC%" QuickFolders.rc

:: ─────────────────────────────────────────────────────────────────────────
::   Compile EXE
:: ─────────────────────────────────────────────────────────────────────────

IF EXIST "%OUT%" (
	CALL :DISP_MSG "Cleaning up existing %OUT%..." 0
	TASKKILL /F /IM QuickFolders.exe >NUL 2>&1
    DEL /F /Q "%OUT%"
)

CALL :DISP_MSG "Compiling %OUT%..." 0

"%CSC%" /unsafe /nologo /target:winexe /platform:x86 /optimize+ /debug- /nowarn:1591 /filealign:512 /win32res:%RES% /main:Program /out:%OUT% /reference:%VB% ^
%RESOURCES% ^
%SRC%

IF NOT EXIST "%OUT%" (
	CALL :ERROR_MESSAGE_EXIT "Build failed." 20
)

FOR %%F IN (%OUT%) DO CALL :DISP_MSG "Build complete. Final file size of %OUT%: %%~zF bytes" 0

:: ─────────────────────────────────────────────────────────────────────────
::   Inject application manifest
:: ─────────────────────────────────────────────────────────────────────────

CALL :DISP_MSG "Injecting application manifest to make application DPI aware." 0

:: Wait until QuickFolders.exe is not locked
:WAIT_FOR_UNLOCK
powershell -Command ^
  "$f='%OUT%'; try { $s = [System.IO.File]::Open($f, 'Open', 'ReadWrite', 'None'); $s.Close(); exit 0 } catch { exit 1 }"
IF %ERRORLEVEL% NEQ 0 (
    TIMEOUT /T 1 >NUL
    GOTO :WAIT_FOR_UNLOCK
)

"%MT%" -manifest app.manifest -outputresource:%OUT%;#1
IF %ERRORLEVEL% NEQ 0 (
    CALL :ERROR_MESSAGE_EXIT "mt.exe failed to inject manifest." 22
)

:: ─────────────────────────────────────────────────────────────────────────
::   Sign EXE
:: ─────────────────────────────────────────────────────────────────────────

CALL :DISP_MSG "Signing %OUT% with self-signed certificate..." 0

SET CERT_PATH=%SCRIPT_DIR%\QuickFoldersDevCert.pfx
CALL "%SCRIPT_DIR%\cert_pwd.bat"

"%SIGN_TOOL%" sign /f "%CERT_PATH%" /p %CERT_PWD% /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 %OUT%

:: ─────────────────────────────────────────────────────────────────────────
::   Create setup EXE
:: ─────────────────────────────────────────────────────────────────────────

IF EXIST "%SETUP_EXE%" (
    ECHO Cleaning up existing %SETUP_EXE%...
    DEL /F /Q "%SETUP_EXE%"
)

CALL :DISP_MSG "Creating installer..." 0

"%MAKE_NSIS%" installer.nsi

IF NOT EXIST "%SETUP_EXE%" (
    CALL :ERROR_MESSAGE_EXIT "Failed to generate installer %SETUP_EXE%." 23
)

:: ─────────────────────────────────────────────────────────────────────────
::   Create checksum file
:: ─────────────────────────────────────────────────────────────────────────

CALL :DISP_MSG "Creating checksum file %SHA_OUT%..." 0

IF EXIST "%SHA_OUT%" (
    ECHO Cleaning up existing %SHA_OUT%...
    DEL /F /Q "%SHA_OUT%"
)

FOR %%F IN (*.exe) DO (
    FOR /F %%H IN ('CertUtil -hashfile "%%F" SHA512 ^| FIND /V ":" ^| FIND /V "CertUtil"') DO (
        SET "HASH=%%H"
    )
    >> %SHA_OUT% ECHO !HASH!  %%F
)

CALL :DISP_MSG "SHA-512 checksums written to %SHA_OUT%" 0

:: ─────────────────────────────────────────────────────────────────────────
::   Create checksum file
:: ─────────────────────────────────────────────────────────────────────────

CALL :DISP_MSG "Running installer %SETUP_EXE%..." 2

START "" "%SETUP_EXE%"

ENDLOCAL
EXIT /B 0

:: ─────────────────────────────────────────────────────────────────────────
::   Functions
:: ─────────────────────────────────────────────────────────────────────────

:ERROR_MESSAGE_EXIT
	SET MSG=%1
	SET MSG=%MSG:~1,-1%
	SET MSG=%DATE% %TIME% - %MSG%
	SET CODE=%2
	COLOR 4F
	ECHO.
	ECHO   CODE:  %CODE%
	ECHO   ERROR: !MSG!
	ECHO   Press any key to end...
	PAUSE >NUL
	EXIT %CODE%

:DISP_MSG
	SET MSG=%1
	SET MSG=%MSG:~1,-1%
    SET MSG=%DATE% %TIME% - %MSG%
	SET /A DELAY_SEC=%2+1
	ECHO.
	ECHO   !MSG!
	TIMEOUT /T %DELAY_SEC% /NOBREAK >NUL
	GOTO :EOF
