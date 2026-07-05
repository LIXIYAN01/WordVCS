@echo off
REM ==============================================
REM  WordVCS COM Add-in 卸载脚本
REM ==============================================

set DLL_PATH=%~dp0bin\Release\net48\WordVCS.AddIn.dll
set REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

echo.
echo WordVCS - 卸载 COM Add-in
echo ========================================

REM 1. Remove registry keys
echo [1/3] 删除注册表项...
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /f >nul 2>&1
echo   已删除

REM 2. Unregister COM
echo [2/3] 卸载 COM 组件...
"%REGASM%" "%DLL_PATH%" /unregister
echo   已卸载

echo [3/3] 卸载完成！
echo.
pause
