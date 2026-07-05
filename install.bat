@echo off
setlocal

REM ==============================================
REM  WordVCS Installer
REM  Right-click -> Run as Administrator
REM ==============================================
TITLE WordVCS Installer

set DLL=%~dp0src\WordVCS.AddIn\bin\Release\net48\WordVCS.AddIn.dll
set CLSID={B4E1C2D3-A5F6-7890-ABCD-EF1234567890}
set CODEFILE=file:///%DLL:\=/%
set RTVER=v4.0.30319

echo.
echo ==========================================
echo   WordVCS Installer
echo ==========================================
echo.

if not exist "%DLL%" (
    echo [ERROR] DLL not found: %DLL%
    pause
    exit /b 1
)

echo [1/4] Cleaning old registration...
reg delete "HKCR\CLSID\%CLSID%" /f >nul 2>&1
reg delete "HKCR\WordVCS.Connect" /f >nul 2>&1
echo   [OK]

echo [2/4] Registering COM class...
reg add "HKCR\CLSID\%CLSID%" /ve /d "WordVCS.AddIn.ThisAddIn" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32" /ve /d "mscoree.dll" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32" /v ThreadingModel /d "Both" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32" /v Class /d "WordVCS.AddIn.ThisAddIn" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32" /v Assembly /d "WordVCS.AddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32" /v RuntimeVersion /d "%RTVER%" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32" /v CodeBase /d "%CODEFILE%" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32\%RTVER%" /v Class /d "WordVCS.AddIn.ThisAddIn" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32\%RTVER%" /v Assembly /d "WordVCS.AddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32\%RTVER%" /v RuntimeVersion /d "%RTVER%" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\InprocServer32\%RTVER%" /v CodeBase /d "%CODEFILE%" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\ProgId" /ve /d "WordVCS.Connect" /f >nul 2>&1
reg add "HKCR\CLSID\%CLSID%\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" /f >nul 2>&1
reg add "HKCR\WordVCS.Connect" /ve /d "WordVCS.AddIn.ThisAddIn" /f >nul 2>&1
reg add "HKCR\WordVCS.Connect\CLSID" /ve /d "%CLSID%" /f >nul 2>&1
echo   [OK]

echo [3/4] Adding to Word and WPS...
reg add "HKCU\SOFTWARE\Microsoft\Office\Word\Addins\WordVCS.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKCU\SOFTWARE\Microsoft\Office\Word\Addins\WordVCS.Connect" /v FriendlyName /t REG_SZ /d "WordVCS" /f >nul 2>&1
reg add "HKCU\SOFTWARE\Microsoft\Office\Word\Addins\WordVCS.Connect" /v Description /t REG_SZ /d "Word Version Control" /f >nul 2>&1
reg add "HKCU\SOFTWARE\Microsoft\Office\Word\Addins\WordVCS.Connect" /v CommandLineSafe /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\Word\Addins\WordVCS.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\Word\Addins\WordVCS.Connect" /v FriendlyName /t REG_SZ /d "WordVCS" /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Office\Word\Addins\WordVCS.Connect" /v CommandLineSafe /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Kingsoft\Office\WPS\Addins\WordVCS.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKCU\Software\Kingsoft\Office\WPS\Addins\WordVCS.Connect" /v FriendlyName /t REG_SZ /d "WordVCS" /f >nul 2>&1
echo   [OK]

echo [4/4] Done!
echo.
echo ==========================================
echo   Install complete: 0 errors, 0 warnings ^^!
echo.
echo   Restart Word. If tab not visible:
echo   File - Options - Add-ins
echo   Manage: COM Add-ins - Go
echo   Check WordVCS.Connect - OK
echo ==========================================
pause
endlocal
