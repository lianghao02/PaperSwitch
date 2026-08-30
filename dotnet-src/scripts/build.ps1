# ============================================================
# PaperSwitch 一鍵編譯發行腳本 (Build & Publish Script)
# ============================================================
param (
    [switch]$SelfContained = $false,
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

# 自動偵測 dotnet SDK 路徑
if (Test-Path "$env:LOCALAPPDATA\Microsoft\dotnet\dotnet.exe") {
    $env:DOTNET_ROOT = "$env:LOCALAPPDATA\Microsoft\dotnet"
    $env:PATH = "$env:LOCALAPPDATA\Microsoft\dotnet;$env:PATH"
    $DotnetExe = "$env:LOCALAPPDATA\Microsoft\dotnet\dotnet.exe"
} elseif (Test-Path "$env:USERPROFILE\.dotnet\dotnet.exe") {
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
    "-o", $PublishDir
)

if ($SelfContained) {
    Write-Host "[PaperSwitch] Mode: Self-Contained" -ForegroundColor Yellow
    $PublishArgs += @(
        "--self-contained", "true",
        "-p:PublishSingleFile=true",
        "-p:IncludeNativeLibrariesForSelfExtract=true",
        "-p:EnableCompressionInSingleFile=true"
    )
} else {
    Write-Host "[PaperSwitch] Mode: Framework-Dependent" -ForegroundColor Yellow
    $PublishArgs += @(
        "--self-contained", "false",
        "-p:PublishSingleFile=true",
        "-p:IncludeNativeLibrariesForSelfExtract=true"
    )
}

& $DotnetExe @PublishArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "[PaperSwitch] Build & Publish Successful!" -ForegroundColor Green
    $ExePath = Join-Path $PublishDir "PaperSwitch.exe"
    if (Test-Path $ExePath) {
        $Size = [Math]::Round((Get-Item $ExePath).Length / 1MB, 2)
        Write-Host "   Output: $ExePath ($Size MB)" -ForegroundColor Green
    }
    Write-Host "============================================================" -ForegroundColor Green
} else {
    Write-Host "[PaperSwitch] Build Failed, ExitCode: $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}
