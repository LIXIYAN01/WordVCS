@echo off
setlocal

REM ==============================================
REM  WordVCS Uninstaller
REM ==============================================
TITLE WordVCS Uninstaller

set DLL=%~dp0src\WordVCS.AddIn\bin\Release\net48\WordVCS.AddIn.dll
set REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

echo.
echo ==========================================
echo   WordVCS Uninstaller
echo ==========================================
echo.

echo [1/3] Removing Word add-in...
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /f >nul 2>&1
reg delete "HKLM\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /f >nul 2>&1

echo [2/3] Removing WPS add-in...
reg delete "HKCU\Software\Kingsoft\Office\WPS\Addins\WordVCS.Connect" /f >nul 2>&1
reg delete "HKLM\Software\Kingsoft\Office\WPS\Addins\WordVCS.Connect" /f >nul 2>&1

echo [3/3] Unregistering COM...
call "%REGASM%" "%DLL%" /unregister /tlb /nologo >nul 2>&1

echo.
echo ==========================================
echo   Uninstall complete.
echo ==========================================
pause
endlocal
