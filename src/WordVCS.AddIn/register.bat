@echo off
REM ==============================================
REM  WordVCS COM Add-in 注册脚本
REM  以管理员身份运行此脚本
REM ==============================================

set DLL_PATH=%~dp0bin\Release\net48\WordVCS.AddIn.dll
set REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

echo.
echo WordVCS - 注册 COM Add-in
echo ========================================
echo.

if not exist "%DLL_PATH%" (
    echo 错误: 找不到 DLL 文件
    echo 路径: %DLL_PATH%
    echo 请先运行 build.ps1 编译项目
    pause
    exit /b 1
)

REM 1. Register managed DLL for COM
echo [1/3] 注册 COM 组件...
"%REGASM%" "%DLL_PATH%" /codebase /tlb
if %ERRORLEVEL% NEQ 0 (
    echo 注册失败！请以管理员身份运行此脚本。
    pause
    exit /b 1
)
echo   注册成功

REM 2. Add Word Add-in registry keys
echo [2/3] 添加 Word 插件注册表项...
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v FriendlyName /t REG_SZ /d "WordVCS - 论文版本控制" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v Description /t REG_SZ /d "Word 文档版本控制系统 VSTO 插件" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\Word\Addins\WordVCS.Connect" /v CommandLineSafe /t REG_DWORD /d 1 /f >nul 2>&1
echo   注册表已更新

REM 3. Done
echo [3/3] 注册完成！
echo.
echo ========================================
echo   WordVCS 已成功注册！
echo   请重新启动 Word，在"开始"旁查看"论文版本"标签
echo ========================================
echo.
pause
