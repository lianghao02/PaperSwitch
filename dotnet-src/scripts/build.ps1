# ============================================================
# PaperSwitch 一鍵編譯發行腳本 (Build & Publish Script)
# ============================================================
param (
    [switch]$SelfContained = $false,
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

# 自動偵測 dotnet SDK 路徑
if (Test-Path "$env:USERPROFILE\.dotnet\dotnet.exe") {
    $env:DOTNET_ROOT = "$env:USERPROFILE\.dotnet"
    $env:PATH = "$env:USERPROFILE\.dotnet;$env:PATH"
    $DotnetExe = "$env:USERPROFILE\.dotnet\dotnet.exe"
} else {
    $DotnetExe = "dotnet"
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Resolve-Path (Join-Path $ScriptDir "..\..")
$SrcDir = Join-Path $ProjectRoot "dotnet-src\src\PaperSwitch"
$DistDir = Join-Path $ProjectRoot "dist"
$PublishDir = Join-Path $DistDir "publish"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "🚀 [PaperSwitch] 啟動 C# 12 / .NET 8 WPF 應用程式建置" -ForegroundColor Cyan
Write-Host "   專案目錄: $SrcDir"
Write-Host "   發行目錄: $PublishDir"
Write-Host "   組態:     $Configuration"
Write-Host "   自包含:   $SelfContained"
Write-Host "============================================================" -ForegroundColor Cyan

if (Test-Path $PublishDir) {
    Remove-Item -Path $PublishDir -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -ItemType Directory -Path $PublishDir -Force | Out-Null

$PublishArgs = @(
    "publish",
    "$SrcDir\PaperSwitch.csproj",
    "-c", $Configuration,
    "-r", "win-x64",
    "--no-restore",
    "-o", $PublishDir
)

if ($SelfContained) {
    Write-Host "📦 發行模式: Self-Contained (單一免安裝 Exe，內嵌 .NET 8 執行階段)..." -ForegroundColor Yellow
    $PublishArgs += @(
        "--self-contained", "true",
        "-p:PublishSingleFile=true",
        "-p:IncludeNativeLibrariesForSelfExtract=true",
        "-p:EnableCompressionInSingleFile=true"
    )
} else {
    Write-Host "⚡ 發行模式: Framework-Dependent (極致輕量秒開，複用本機 .NET 8 執行階段)..." -ForegroundColor Yellow
    $PublishArgs += @(
        "--self-contained", "false",
        "-p:PublishSingleFile=true",
        "-p:IncludeNativeLibrariesForSelfExtract=true"
    )
}

& $DotnetExe @PublishArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "✅ [PaperSwitch] 編譯發行成功！" -ForegroundColor Green
    $ExePath = Join-Path $PublishDir "PaperSwitch.exe"
    if (Test-Path $ExePath) {
        $Size = (Get-Item $ExePath).Length / 1MB
        $Msg = ("   產生成品: {0} ({1:N2} MB)" -f $ExePath, $Size)
        Write-Host $Msg -ForegroundColor Green
    }
    Write-Host "============================================================" -ForegroundColor Green
} else {
    Write-Host "❌ [PaperSwitch] 編譯發行失敗，結束代碼: $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}
