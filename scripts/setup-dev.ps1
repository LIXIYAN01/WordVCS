<#
.SYNOPSIS
    WordVCS 开发环境搭建脚本
#>

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  WordVCS 开发环境设置" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Check .NET Framework
Write-Host "`n检查 .NET Framework 4.8..." -ForegroundColor Yellow
$net48 = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
if ($net48.Release -ge 528040) {
    Write-Host "  .NET Framework 4.8 已安装" -ForegroundColor Green
} else {
    Write-Host "  警告: 需要 .NET Framework 4.8" -ForegroundColor Red
}

# Check Visual Studio
Write-Host "`n检查 Visual Studio 2022..." -ForegroundColor Yellow
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $vsPath = & $vswhere -latest -property installationPath
    Write-Host "  Visual Studio: $vsPath" -ForegroundColor Green
} else {
    Write-Host "  警告: 未检测到 Visual Studio" -ForegroundColor Red
    Write-Host "  请安装 Visual Studio 2022 Community (免费)" -ForegroundColor Yellow
    Write-Host "  下载: https://visualstudio.microsoft.com/zh-hans/downloads/" -ForegroundColor Yellow
}

# Check VSTO workload
Write-Host "`n检查 VSTO 开发负载..." -ForegroundColor Yellow
$vstoPath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Shared\VSTO"
if (Test-Path $vstoPath) {
    Write-Host "  VSTO SDK 已安装" -ForegroundColor Green
} else {
    Write-Host "  警告: VSTO 开发负载未安装" -ForegroundColor Red
    Write-Host "  在 Visual Studio Installer 中勾选 'Office/SharePoint 开发'" -ForegroundColor Yellow
}

# Check Office
Write-Host "`n检查 Microsoft Word..." -ForegroundColor Yellow
$wordPath = "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE"
if (Test-Path $wordPath) {
    Write-Host "  Word 已安装: $wordPath" -ForegroundColor Green
} else {
    Write-Host "  警告: 未在默认位置找到 Word" -ForegroundColor Red
    Write-Host "  需要 Microsoft Word 2016/2019/2021 或 Microsoft 365 桌面版" -ForegroundColor Yellow
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  环境检查完成" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
