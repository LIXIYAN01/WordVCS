<#
.SYNOPSIS
    WordVCS 发布打包脚本
.DESCRIPTION
    编译 Release 版本并创建可分发的 ZIP 包和 VSTO 安装文件。
    需要在安装了 VSTO 开发负载的 Visual Studio 2022 中运行。
#>

param(
    [string]$Configuration = "Release",
    [string]$OutputDir = ".\publish"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SolutionDir = Resolve-Path "$ScriptDir\.."
$SolutionFile = "$SolutionDir\WordVCS.sln"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  WordVCS 发布打包" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Step 1: Build
Write-Host "`n[1/3] 编译 Release 版本..." -ForegroundColor Yellow
& "$ScriptDir\build.ps1" -Configuration $Configuration

# Step 2: Collect outputs
Write-Host "`n[2/3] 收集构建产物..." -ForegroundColor Yellow
$publishDir = Join-Path $SolutionDir $OutputDir
if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }
New-Item -ItemType Directory -Path $publishDir -Force | Out-Null

$buildOutput = "$SolutionDir\src\WordVCS.AddIn\bin\$Configuration\net48"
if (Test-Path $buildOutput) {
    Copy-Item "$buildOutput\*" $publishDir -Recurse -Force
}

# Copy README
Copy-Item "$SolutionDir\README.md" $publishDir -Force

# Step 3: Create ZIP
Write-Host "`n[3/3] 创建 ZIP 包..." -ForegroundColor Yellow
$zipPath = "$SolutionDir\WordVCS-Plugin.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($publishDir, $zipPath)

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  打包完成!" -ForegroundColor Green
Write-Host "  输出: $zipPath" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
