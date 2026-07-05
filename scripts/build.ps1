<#
.SYNOPSIS
    WordVCS 项目构建脚本
.DESCRIPTION
    还原 NuGet 包并编译整个解决方案。
    需要 Visual Studio 2022 并安装了 Office/SharePoint 开发负载。
.PARAMETER Configuration
    Debug 或 Release，默认 Release
.EXAMPLE
    .\build.ps1
    .\build.ps1 -Configuration Debug
#>

param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SolutionDir = Resolve-Path "$ScriptDir\.."
$SolutionFile = "$SolutionDir\WordVCS.sln"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  WordVCS 构建脚本" -ForegroundColor Cyan
Write-Host "  配置: $Configuration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Step 1: Check prerequisites
Write-Host "`n[1/4] 检查构建环境..." -ForegroundColor Yellow

# Find MSBuild
$msbuild = $null
$vsPath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022"
if (Test-Path "$vsPath\BuildTools\MSBuild\Current\Bin\MSBuild.exe") {
    $msbuild = "$vsPath\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
} elseif (Test-Path "$vsPath\Community\MSBuild\Current\Bin\MSBuild.exe") {
    $msbuild = "$vsPath\Community\MSBuild\Current\Bin\MSBuild.exe"
} elseif (Test-Path "$vsPath\Professional\MSBuild\Current\Bin\MSBuild.exe") {
    $msbuild = "$vsPath\Professional\MSBuild\Current\Bin\MSBuild.exe"
} elseif (Test-Path "$vsPath\Enterprise\MSBuild\Current\Bin\MSBuild.exe") {
    $msbuild = "$vsPath\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
}

if (-not $msbuild -or -not (Test-Path $msbuild)) {
    Write-Host "错误: 未找到 MSBuild。请安装 Visual Studio 2022。" -ForegroundColor Red
    Write-Host "需要安装以下负载:" -ForegroundColor Red
    Write-Host "  - Office/SharePoint 开发" -ForegroundColor Red
    Write-Host "  - .NET 桌面开发" -ForegroundColor Red
    exit 1
}
Write-Host "  MSBuild: $msbuild" -ForegroundColor Green

# Step 2: Restore NuGet packages
Write-Host "`n[2/4] 还原 NuGet 包..." -ForegroundColor Yellow
Push-Location $SolutionDir
try {
    dotnet restore WordVCS.sln
    Write-Host "  NuGet 包还原完成" -ForegroundColor Green
} finally {
    Pop-Location
}

# Step 3: Build
Write-Host "`n[3/4] 编译解决方案..." -ForegroundColor Yellow
Push-Location $SolutionDir
try {
    & $msbuild WordVCS.sln /p:Configuration=$Configuration /m /v:minimal
    if ($LASTEXITCODE -ne 0) {
        Write-Host "编译失败！请检查错误信息。" -ForegroundColor Red
        exit $LASTEXITCODE
    }
    Write-Host "  编译成功！" -ForegroundColor Green
} finally {
    Pop-Location
}

# Step 4: Summary
Write-Host "`n[4/4] 构建完成" -ForegroundColor Yellow
Write-Host "输出文件位置:" -ForegroundColor White
$outputDir = "$SolutionDir\src\WordVCS.AddIn\bin\$Configuration\net48"
if (Test-Path $outputDir) {
    Get-ChildItem $outputDir -Filter "*.dll" | ForEach-Object {
        Write-Host "  $($_.FullName)" -ForegroundColor Gray
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  构建完成！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "提示: 在 Visual Studio 中按 F5 启动调试" -ForegroundColor White
Write-Host "      或直接打开 Word 加载插件" -ForegroundColor White
