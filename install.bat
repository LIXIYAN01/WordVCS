@echo off
setlocal EnableDelayedExpansion

REM ==============================================
REM  WordVCS Installer
REM  Right-click -> Run as Administrator
REM ==============================================
TITLE WordVCS Installer

set DLL=%~dp0src\WordVCS.AddIn\bin\Release\net48\WordVCS.AddIn.dll
set REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe
set CLSID={B4E1C2D3-A5F6-7890-ABCD-EF1234567890}

echo.
echo ==========================================
echo   WordVCS Installer
echo ==========================================
echo.

if not exist "%DLL%" (
    echo [ERROR] DLL not found:
    echo   %DLL%
    echo Please run this script from the WordVCS
    echo project root directory.
    pause
    exit /b 1
)

REM --- Step 1: COM Registration ---
echo [1/4] Registering COM component...
call "%REGASM%" "%DLL%" /codebase /tlb /nologo >nul 2>&1
if !ERRORLEVEL! NEQ 0 (
    echo [FAIL] Run as Administrator!
    pause
    exit /b 1
)

REM --- Step 2: Office Add-in Category ---
echo [2/4] Adding Office Add-in category...
reg add "HKCR\CLSID\%CLSID%\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\Programmable" /f >nul 2>&1

REM --- Step 3: Word Add-in Registry ---
echo [3/4] Registering with Word...

REM HKCU (user-level)
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v FriendlyName /t REG_SZ /d "WordVCS - LunWen BanBen KongZhi" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v Description /t REG_SZ /d "Word VCS for thesis writing" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v CommandLineSafe /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v Manifest /t REG_SZ /d "file:///%DLL:\=/%" /f >nul 2>&1

REM HKLM (machine-level - some Office builds need this)
reg add "HKLM\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKLM\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v FriendlyName /t REG_SZ /d "WordVCS - LunWen BanBen KongZhi" /f >nul 2>&1
reg add "HKLM\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v Description /t REG_SZ /d "Word VCS for thesis writing" /f >nul 2>&1
reg add "HKLM\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v CommandLineSafe /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v Manifest /t REG_SZ /d "file:///%DLL:\=/%" /f >nul 2>&1

REM WPS Office registration (same COM add-in model)
reg add "HKCU\Software\Kingsoft\Office\WPS\Addins\WordVCS.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKCU\Software\Kingsoft\Office\WPS\Addins\WordVCS.Connect" /v FriendlyName /t REG_SZ /d "WordVCS" /f >nul 2>&1
reg add "HKLM\Software\Kingsoft\Office\WPS\Addins\WordVCS.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKLM\Software\Kingsoft\Office\WPS\Addins\WordVCS.Connect" /v FriendlyName /t REG_SZ /d "WordVCS" /f >nul 2>&1

echo [ OK ]

REM --- Step 4: Done ---
echo [4/4] Done!
echo.
echo ==========================================
echo   Installation complete!
echo.
echo   1. Restart Word
echo   2. Look for the tab next to Home
echo   3. If not visible: File - Options -
echo      Add-ins - Manage: COM Add-ins -
echo      Go... - check WordVCS.Connect
echo ==========================================
echo.
pause
endlocal
