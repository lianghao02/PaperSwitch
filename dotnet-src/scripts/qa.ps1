# ============================================================
# PaperSwitch 品質檢核與自動化驗證腳本 (QA Script)
# ============================================================
$ErrorActionPreference = "Stop"

if (Test-Path "$env:USERPROFILE\.dotnet\dotnet.exe") {
    $env:DOTNET_ROOT = "$env:USERPROFILE\.dotnet"
    $env:PATH = "$env:USERPROFILE\.dotnet;$env:PATH"
    $DotnetExe = "$env:USERPROFILE\.dotnet\dotnet.exe"
} else {
    $DotnetExe = "dotnet"
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SlnPath = Resolve-Path (Join-Path $ScriptDir "..\PaperSwitch.sln")

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "🔍 [QA] 開始執行 PaperSwitch 建置與單元測試檢驗..." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

Write-Host "1. 還原與清理建置..." -ForegroundColor Yellow
& $DotnetExe clean $SlnPath
& $DotnetExe build $SlnPath -c Release

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ 建置失敗！" -ForegroundColor Red
    exit 1
}

Write-Host "2. 執行單元測試..." -ForegroundColor Yellow
& $DotnetExe test $SlnPath -c Release --verbosity normal

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ 單元測試失敗！" -ForegroundColor Red
    exit 1
}

Write-Host "============================================================" -ForegroundColor Green
Write-Host "🎉 [QA] 全部建置與測試檢驗通過！" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
